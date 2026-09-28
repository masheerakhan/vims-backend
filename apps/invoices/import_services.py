import io
import re
import csv
from uuid import uuid4
from decimal import Decimal, InvalidOperation
from typing import List, Dict, Any, Optional

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

from django.db import transaction, models
from django.db.models import Sum
from django.utils import timezone
from django.utils.text import slugify

from apps.core.models import ScopeNode, Organization
from apps.vendors.models import Vendor, OperationalStatus, MarketingStatus
from apps.budgets.models import (
    Budget,
    BudgetLine,
    BudgetCategory,
    BudgetSubCategory,
    BudgetStatus,
    SourceType,
)
from apps.budgets.services import consume_budget_line_direct
from apps.invoices.models import (
    Invoice,
    InvoiceStatus,
    InvoiceEntrySource,
    InvoiceAllocation,
    InvoiceAllocationStatus,
    InvoiceAllocationSource,
    InvoiceImportBatch,
    InvoiceImportRow,
    InvoiceImportBatchStatus,
    InvoiceImportRowStatus,
)

COLUMN_MAPPING = {
    "invoice_number": [
        "invoice no.", "invoice no", "invoice number", "invoice_number",
        "invoice_no", "inv no.", "inv no", "inv_no", "invoice #", "bill no"
    ],
    "vendor_name": [
        "vendor", "vendor name", "vendor_name", "company name",
        "supplier", "supplier name", "party name"
    ],
    "work_description": [
        "work description", "work_description", "description",
        "particulars", "narration", "details", "item description"
    ],
    "amount": [
        "amount", "amount (inr)", "amount inr", "total amount",
        "invoice amount", "value", "inr"
    ],
    "main_head": [
        "main head", "main_head", "head type", "head", "type", "region"
    ],
    "budget_head": [
        "budget head", "budget_head", "category", "budget category",
        "park", "park name", "park / head"
    ],
    "sub_category": [
        "sub category", "sub_category", "subcategory",
        "sub-category", "activity", "line item"
    ],
}


def _clean_amount(val: Any) -> Optional[Decimal]:
    if val is None:
        return None
    s = str(val).strip()
    if not s:
        return None
    # Strip currency symbols, commas, spaces
    s = re.sub(r"[?\$,\s]", "", s)
    try:
        amt = Decimal(s)
        return amt if amt > 0 else None
    except (InvalidOperation, ValueError):
        return None


def _normalize_text(s: str) -> str:
    if not s:
        return ""
    s = str(s).strip().lower()
    s = re.sub(r"[\*\_]", " ", s)
    return re.sub(r"\s+", " ", s).strip()


def generate_invoice_import_template() -> io.BytesIO:
    """
    Generate a styled Excel template matching the required 7 columns:
    Invoice no. | Vendor | Work Description | Amount | Main Head | Budget Head | Sub Category
    """
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Invoice Import"
    ws.views.sheetView[0].showGridLines = True

    headers = [
        ("Invoice no.", 18),
        ("Vendor", 30),
        ("Work Description", 45),
        ("Amount", 16),
        ("Main Head", 16),
        ("Budget Head", 26),
        ("Sub Category", 34),
    ]

    header_font = Font(name="Segoe UI", size=11, bold=True, color="FFFFFF")
    header_fill = PatternFill(start_color="EA580C", end_color="EA580C", fill_type="solid")
    header_align = Alignment(horizontal="center", vertical="center", wrap_text=True)

    thin_border = Border(
        left=Side(style="thin", color="CBD5E1"),
        right=Side(style="thin", color="CBD5E1"),
        top=Side(style="thin", color="CBD5E1"),
        bottom=Side(style="thin", color="CBD5E1"),
    )

    ws.row_dimensions[1].height = 30

    for col_idx, (header_text, col_width) in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col_idx, value=header_text)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = header_align
        cell.border = thin_border
        col_letter = get_column_letter(col_idx)
        ws.column_dimensions[col_letter].width = col_width

    # Sample rows demonstrating Corporate, Park, and Direct Card spends
    sample_rows = [
        [
            "ID/26-27/00060",
            "Identica",
            "WORKS AT YOGESHWAR DUTT ACADEMY Artwork",
            106750,
            "Corporate",
            "Others",
            "Miscellaneous",
        ],
        [
            "ID/26-27/00016",
            "Identica",
            "Hosur Flex Hoarding",
            72000,
            "Park",
            "Hosur",
            "Onsite Hoarding perimeter branding",
        ],
        [
            "INV-27818",
            "VSol4U Technologies Pvt. Ltd",
            "Sahibabad CWC Data Capture (Shoot) - Dated: 13th May 2026",
            25000,
            "Park",
            "Sahibabad 2",
            "CMVs, Panorama & Timelapse",
        ],
        [
            "SHPL/PI/01/26-27",
            "Saffron Habitat Pvt. Ltd.",
            "Coffee Table book Designing",
            500000,
            "Corporate",
            "Content Marketing & Assets",
            "Corporate Brochure",
        ],
        [
            "",
            "Taruna-CC",
            "Google Ads & Youtube",
            58699,
            "Corporate",
            "Content Marketing & Assets",
            "Digital Media Buying",
        ],
    ]

    data_font = Font(name="Segoe UI", size=10)
    data_align = Alignment(vertical="center")

    for row_idx, row_data in enumerate(sample_rows, 2):
        ws.row_dimensions[row_idx].height = 22
        for col_idx, val in enumerate(row_data, 1):
            cell = ws.cell(row=row_idx, column=col_idx, value=val)
            cell.font = data_font
            cell.alignment = data_align
            cell.border = thin_border

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output


def parse_invoice_import_file(uploaded_file) -> List[Dict[str, Any]]:
    """
    Parse an uploaded .xlsx, .xls, or .csv file into structured row dicts.
    """
    name = uploaded_file.name.lower()
    rows = []

    if name.endswith(".csv"):
        uploaded_file.seek(0)
        content = uploaded_file.read()
        try:
            text = content.decode("utf-8-sig")
        except UnicodeDecodeError:
            text = content.decode("latin-1")
        reader = csv.reader(io.StringIO(text))
        raw_headers = next(reader, None)
        if not raw_headers:
            return []

        header_map = {}
        for col_idx, raw_h in enumerate(raw_headers):
            norm_h = _normalize_text(raw_h)
            for field, aliases in COLUMN_MAPPING.items():
                if norm_h in aliases or any(_normalize_text(a) == norm_h for a in aliases):
                    header_map[col_idx] = field
                    break

        for row_idx, raw_row in enumerate(reader, 2):
            if not any(str(c).strip() for c in raw_row):
                continue
            row_dict = {}
            for col_idx, val in enumerate(raw_row):
                field_name = header_map.get(col_idx)
                if field_name:
                    row_dict[field_name] = str(val).strip() if val is not None else ""
            row_dict["_row_number"] = row_idx
            rows.append(row_dict)

    else:
        uploaded_file.seek(0)
        wb = openpyxl.load_workbook(uploaded_file, data_only=True)
        ws = wb.active
        raw_headers = [cell.value for cell in ws[1]]
        if not raw_headers:
            return []

        header_map = {}
        for col_idx, raw_h in enumerate(raw_headers):
            norm_h = _normalize_text(raw_h)
            for field, aliases in COLUMN_MAPPING.items():
                if norm_h in aliases or any(_normalize_text(a) == norm_h for a in aliases):
                    header_map[col_idx] = field
                    break

        for row_idx, sheet_row in enumerate(ws.iter_rows(min_row=2, values_only=True), 2):
            if not any(sheet_row):
                continue
            row_dict = {}
            for col_idx, val in enumerate(sheet_row):
                field_name = header_map.get(col_idx)
                if field_name:
                    row_dict[field_name] = str(val).strip() if val is not None else ""
            row_dict["_row_number"] = row_idx
            rows.append(row_dict)

    return rows


def append_batch_activity_log(
    batch: InvoiceImportBatch,
    event: str,
    message: str,
    user=None,
    extra_details: Optional[Dict[str, Any]] = None,
):
    """Appends an immutable audit/activity log entry to the batch record."""
    logs = list(batch.activity_logs or [])
    user_display = (
        f"{user.get_full_name() or user.email or user.username}"
        if user and hasattr(user, "email")
        else "System"
    )
    entry = {
        "id": len(logs) + 1,
        "timestamp": timezone.now().isoformat(),
        "event": event,
        "message": message,
        "user": user_display,
        "details": extra_details or {},
    }
    logs.append(entry)
    batch.activity_logs = logs


@transaction.atomic
def create_invoice_import_batch(
    org: Organization,
    file_name: str,
    parsed_rows: List[Dict[str, Any]],
    user=None,
) -> InvoiceImportBatch:
    """
    Persist an InvoiceImportBatch and create individual InvoiceImportRow records.
    """
    batch = InvoiceImportBatch.objects.create(
        org=org,
        file_name=file_name,
        uploaded_by=user,
        status=InvoiceImportBatchStatus.UPLOADED,
        total_rows=len(parsed_rows),
    )

    append_batch_activity_log(
        batch,
        event="UPLOADED",
        message=f"Uploaded {len(parsed_rows)} rows from file '{file_name}'.",
        user=user,
        extra_details={"total_rows": len(parsed_rows), "file_name": file_name},
    )
    batch.save(update_fields=["activity_logs"])

    rows_to_create = []
    for r in parsed_rows:
        row_num = r.get("_row_number", len(rows_to_create) + 2)
        clean_amt = _clean_amount(r.get("amount"))

        row_obj = InvoiceImportRow(
            batch=batch,
            row_number=row_num,
            raw_data=r,
            invoice_number=str(r.get("invoice_number", "")).strip(),
            vendor_name=str(r.get("vendor_name", "")).strip(),
            work_description=str(r.get("work_description", "")).strip(),
            amount=clean_amt,
            main_head=str(r.get("main_head", "")).strip(),
            budget_head=str(r.get("budget_head", "")).strip(),
            sub_category=str(r.get("sub_category", "")).strip(),
            status=InvoiceImportRowStatus.PENDING,
            error_messages=[],
        )
        rows_to_create.append(row_obj)

    InvoiceImportRow.objects.bulk_create(rows_to_create)
    return batch


def _clean_key(s: str) -> str:
    """Removes all whitespace and non-alphanumeric chars for ultra-tolerant matching."""
    return re.sub(r"[^a-zA-Z0-9]", "", str(s or "")).lower()


def _find_matching_vendor(vendor_name: str, vendors_list: List[Vendor]) -> Optional[Vendor]:
    if not vendor_name:
        return None
    v_norm = _normalize_text(vendor_name)
    v_clean = _clean_key(vendor_name)

    # 1. Exact normalized match
    for v in vendors_list:
        if _normalize_text(v.vendor_name) == v_norm:
            return v

    # 2. Clean key match (ignores spaces, casing, punctuation)
    for v in vendors_list:
        if _clean_key(v.vendor_name) == v_clean:
            return v

    # 3. Match base name without parentheses (e.g. "Pooja (Interior designer)" -> "Pooja")
    base_name = re.sub(r"\(.*?\)", "", vendor_name).strip()
    if base_name:
        base_clean = _clean_key(base_name)
        for v in vendors_list:
            if _clean_key(v.vendor_name) == base_clean:
                return v

    # 4. Partial / Substring match for names with 4+ chars
    if len(v_clean) >= 4:
        for v in vendors_list:
            db_clean = _clean_key(v.vendor_name)
            if db_clean and (v_clean == db_clean or v_clean.startswith(db_clean) or db_clean.startswith(v_clean)):
                return v

    return None


def _find_matching_scope_node(name_or_code: str, scope_nodes: List[ScopeNode]) -> Optional[ScopeNode]:
    if not name_or_code:
        return None
    s_norm = _normalize_text(name_or_code)
    s_clean = _clean_key(name_or_code)
    # 1. Exact or normalized match
    for sn in scope_nodes:
        if _normalize_text(sn.name) == s_norm or _normalize_text(sn.code) == s_norm:
            return sn
        if _clean_key(sn.name) == s_clean or _clean_key(sn.code) == s_clean:
            return sn
    # 2. Match base name without parentheses (e.g. "Koka (Vertical)" -> "Koka")
    base_name = re.sub(r"\(.*?\)", "", name_or_code).strip()
    if base_name:
        base_clean = _clean_key(base_name)
        for sn in scope_nodes:
            if _clean_key(sn.name) == base_clean or _clean_key(sn.name).startswith(base_clean):
                return sn
    # 3. Partial word / prefix match
    for sn in scope_nodes:
        db_clean = _clean_key(sn.name)
        if s_clean in db_clean or db_clean.startswith(s_clean):
            return sn
    return None


def _find_matching_budget(scope_node: ScopeNode, budgets: List[Budget]) -> Optional[Budget]:
    # Match budget for this scope node
    for b in budgets:
        if b.scope_node_id == scope_node.id and b.status == BudgetStatus.ACTIVE:
            return b
    # Match any budget for this scope node
    for b in budgets:
        if b.scope_node_id == scope_node.id:
            return b
    return None


def _find_matching_category(name_or_code: str, categories: List[BudgetCategory]) -> Optional[BudgetCategory]:
    if not name_or_code:
        return None
    c_norm = _normalize_text(name_or_code)
    c_clean = _clean_key(name_or_code)
    for c in categories:
        if _normalize_text(c.name) == c_norm or _normalize_text(c.code) == c_norm:
            return c
        if _clean_key(c.name) == c_clean:
            return c
    # Base name match without parentheses (e.g. "Outdoor local" -> "outdoor (local)")
    base_c = re.sub(r"\(.*?\)", "", name_or_code).strip()
    if base_c:
        base_clean = _clean_key(base_c)
        for c in categories:
            if _clean_key(re.sub(r"\(.*?\)", "", c.name)) == base_clean:
                return c
    for c in categories:
        db_clean = _clean_key(c.name)
        if c_clean in db_clean or db_clean in c_clean:
            return c
    return None


def _find_matching_subcategory(
    name_or_code: str,
    category: Optional[BudgetCategory],
    subcategories: List[BudgetSubCategory]
) -> Optional[BudgetSubCategory]:
    if not name_or_code:
        return None
    sc_norm = _normalize_text(name_or_code)
    # Filter by category if provided
    candidate_list = [sc for sc in subcategories if sc.category_id == category.id] if category else subcategories

    for sc in candidate_list:
        if _normalize_text(sc.name) == sc_norm or _normalize_text(sc.code) == sc_norm:
            return sc
    for sc in candidate_list:
        db_norm = _normalize_text(sc.name)
        if sc_norm in db_norm or db_norm in sc_norm:
            return sc
    return None


@transaction.atomic
def validate_invoice_import_batch(batch: InvoiceImportBatch) -> InvoiceImportBatch:
    """
    Validate all rows in an InvoiceImportBatch:
    - Checks vendor existence; flags new vendors for auto-creation.
    - Resolves ScopeNode (Corporate vs Park) and active Budget.
    - Resolves BudgetCategory and BudgetSubCategory.
    - Resolves or flags BudgetLine.
    """
    batch.status = InvoiceImportBatchStatus.VALIDATING
    batch.save(update_fields=["status", "updated_at"])

    rows = list(batch.rows.all().order_by("row_number"))
    
    # Preload entities for org
    vendors = list(Vendor.objects.filter(org=batch.org))
    scope_nodes = list(ScopeNode.objects.filter(org=batch.org))
    budgets = list(Budget.objects.filter(org=batch.org))
    categories = list(BudgetCategory.objects.filter(org=batch.org))
    subcategories = list(BudgetSubCategory.objects.filter(category__org=batch.org))
    budget_lines = list(BudgetLine.objects.filter(budget__org=batch.org))

    corporate_node = _find_matching_scope_node("corporate", scope_nodes)
    if not corporate_node:
        corporate_node = ScopeNode.objects.filter(org=batch.org, parent__isnull=True).first()

    corporate_budget = _find_matching_budget(corporate_node, budgets) if corporate_node else None

    # Preload existing invoices for deduplication check
    existing_invoices = list(
        Invoice.objects.filter(scope_node__org=batch.org)
        .exclude(status=InvoiceStatus.REJECTED)
        .select_related("vendor", "scope_node")
        .prefetch_related("allocations__category")
    )
    existing_by_item: Dict[Tuple[str, str, Optional[int], Optional[int]], Invoice] = {}
    existing_by_vendor_inv: Dict[Tuple[str, str], Invoice] = {}
    existing_by_hist_key: Dict[str, Invoice] = {}
    for inv in existing_invoices:
        if inv.historical_import_key:
            existing_by_hist_key[inv.historical_import_key] = inv
        v_num = _normalize_text(inv.vendor_invoice_number)
        if v_num and inv.vendor:
            v_name_k = _normalize_text(inv.vendor.vendor_name)
            existing_by_vendor_inv[(v_name_k, v_num)] = inv
            # Index by item key (vendor + invoice_number + scope + category)
            # so multiple lines of the same invoice across different categories/parks can be imported
            for alloc in inv.allocations.all():
                existing_by_item[(v_name_k, v_num, inv.scope_node_id, alloc.category_id)] = inv
            if not inv.allocations.exists():
                existing_by_item[(v_name_k, v_num, inv.scope_node_id, None)] = inv

    valid_count = 0
    error_count = 0
    skipped_count = 0
    new_vendor_count = 0
    batch_errors = []

    seen_new_vendors = set()

    for row in rows:
        # Check if already created or existing for this batch row
        if row.created_invoice_id:
            row.status = InvoiceImportRowStatus.COMMITTED
            row.error_messages = []
            continue

        v_name = row.vendor_name.strip()
        inv_num_clean = row.invoice_number.strip()
        if not inv_num_clean and v_name:
            inv_num_clean = f"EXP-{slugify(v_name)[:12].upper()}-{row.row_number}"

        if inv_num_clean:
            batch_hist_key = f"bulk:{batch.id}:{row.row_number}:{inv_num_clean}"
            if batch_hist_key in existing_by_hist_key:
                row.created_invoice = existing_by_hist_key[batch_hist_key]
                row.status = InvoiceImportRowStatus.COMMITTED
                row.error_messages = []
                continue

        errors = []

        # 1. Vendor check
        v_name = row.vendor_name.strip()
        if not v_name:
            errors.append("Vendor name is required.")
            matched_vendor = None
            is_new = False
        else:
            matched_vendor = _find_matching_vendor(v_name, vendors)
            if matched_vendor:
                is_new = False
            else:
                is_new = True
                norm_v = _normalize_text(v_name)
                if norm_v not in seen_new_vendors:
                    new_vendor_count += 1
                    seen_new_vendors.add(norm_v)

        row.resolved_vendor = matched_vendor
        row.vendor_is_new = is_new

        # 2. Amount check
        if row.amount is None or row.amount <= 0:
            errors.append("Invalid or missing invoice amount.")

        # 3. Scope & Budget resolution
        main_head_norm = _normalize_text(row.main_head)
        budget_head_raw = row.budget_head.strip()
        sub_cat_raw = row.sub_category.strip()

        resolved_scope = None
        resolved_budget = None
        resolved_cat = None
        resolved_subcat = None
        resolved_line = None

        if main_head_norm in ("corporate", "corp", "ho", "head office"):
            # Corporate path
            resolved_scope = corporate_node
            resolved_budget = corporate_budget
            if not resolved_scope:
                errors.append("Corporate scope node not found.")
            if not resolved_budget:
                errors.append("Active Corporate budget not found.")

            # Category from budget_head
            if budget_head_raw:
                resolved_cat = _find_matching_category(budget_head_raw, categories)
                if not resolved_cat:
                    # Try matching subcategory directly under corporate categories
                    resolved_subcat = _find_matching_subcategory(budget_head_raw, None, subcategories)
                    if resolved_subcat:
                        resolved_cat = resolved_subcat.category
                    else:
                        errors.append(f"Corporate Budget Category '{budget_head_raw}' not found.")
            else:
                errors.append("Budget Head (Category) is required for Corporate invoices.")

            # Subcategory from sub_category if not already resolved
            if sub_cat_raw and resolved_cat and not resolved_subcat:
                resolved_subcat = _find_matching_subcategory(sub_cat_raw, resolved_cat, subcategories)

        else:
            # Park / Region path
            # budget_head contains the Park Name
            if budget_head_raw:
                resolved_scope = _find_matching_scope_node(budget_head_raw, scope_nodes)
                if resolved_scope:
                    resolved_budget = _find_matching_budget(resolved_scope, budgets)
                    if not resolved_budget:
                        errors.append(f"No active budget found for Park '{resolved_scope.name}'.")
                else:
                    errors.append(f"Park / Scope '{budget_head_raw}' not found in organization.")
            else:
                errors.append("Park / Budget Head name is required.")

            # Category from sub_category or budget_head
            if sub_cat_raw:
                resolved_cat = _find_matching_category(sub_cat_raw, categories)
                if not resolved_cat:
                    # Try finding subcategory directly
                    resolved_subcat = _find_matching_subcategory(sub_cat_raw, None, subcategories)
                    if resolved_subcat:
                        resolved_cat = resolved_subcat.category
                    else:
                        errors.append(f"Budget Category '{sub_cat_raw}' not found.")

        # Budget line lookup
        if resolved_budget and resolved_cat:
            for bl in budget_lines:
                if bl.budget_id == resolved_budget.id and bl.category_id == resolved_cat.id:
                    if resolved_subcat:
                        if bl.subcategory_id == resolved_subcat.id:
                            resolved_line = bl
                            break
                    else:
                        resolved_line = bl
                        break

        row.resolved_scope_node = resolved_scope
        row.resolved_budget = resolved_budget
        row.resolved_category = resolved_cat
        row.resolved_subcategory = resolved_subcat
        row.resolved_budget_line = resolved_line

        # 4. Duplicate invoice check (Idempotent import protection)
        inv_num_norm = _normalize_text(row.invoice_number)
        v_name_norm = _normalize_text(row.vendor_name)

        existing_match = None
        if inv_num_norm:
            item_key = (
                v_name_norm,
                inv_num_norm,
                resolved_scope.id if resolved_scope else None,
                resolved_cat.id if resolved_cat else None,
            )
            if item_key in existing_by_item:
                existing_match = existing_by_item[item_key]
            elif (v_name_norm, inv_num_norm, None, None) in existing_by_item:
                existing_match = existing_by_item[(v_name_norm, inv_num_norm, None, None)]

        if existing_match and not errors:
            row.status = InvoiceImportRowStatus.SKIPPED
            row.created_invoice = existing_match
            row.error_messages = [
                f"Already posted in system (Invoice #{existing_match.id}, ₹{existing_match.amount}). Skipped to prevent double deduction."
            ]
            skipped_count += 1
            continue

        if errors:
            row.status = InvoiceImportRowStatus.ERROR
            row.error_messages = errors
            error_count += 1
            batch_errors.append(f"Row {row.row_number}: {'; '.join(errors)}")
        else:
            row.status = InvoiceImportRowStatus.VALID
            row.error_messages = []
            valid_count += 1

    InvoiceImportRow.objects.bulk_update(
        rows,
        [
            "resolved_vendor",
            "vendor_is_new",
            "resolved_scope_node",
            "resolved_budget",
            "resolved_category",
            "resolved_subcategory",
            "resolved_budget_line",
            "created_invoice",
            "status",
            "error_messages",
        ]
    )

    batch.valid_rows = valid_count
    batch.error_rows = error_count
    batch.created_vendors_count = new_vendor_count
    batch.error_summary = batch_errors
    batch.status = InvoiceImportBatchStatus.VALIDATED

    append_batch_activity_log(
        batch,
        event="VALIDATED",
        message=f"Validated batch: {valid_count} valid, {skipped_count} skipped (already posted), {error_count} errors.",
        user=None,
        extra_details={
            "valid_rows": valid_count,
            "skipped_rows": skipped_count,
            "error_rows": error_count,
            "new_vendors_detected": new_vendor_count,
        },
    )

    batch.save(
        update_fields=[
            "valid_rows",
            "error_rows",
            "created_vendors_count",
            "error_summary",
            "activity_logs",
            "status",
            "updated_at",
        ]
    )

    return batch


@transaction.atomic
def commit_invoice_import_batch(batch: InvoiceImportBatch, user=None) -> InvoiceImportBatch:
    """
    Commit all valid rows in the batch:
    1. Auto-creates any missing Vendors with Active & Approved status.
    2. Provisions BudgetLines if necessary.
    3. Creates approved Invoice and InvoiceAllocation records.
    4. Atomically consumes amounts from BudgetLines and logs BudgetConsumption.
    """
    if batch.status != InvoiceImportBatchStatus.VALIDATED:
        batch = validate_invoice_import_batch(batch)

    batch.status = InvoiceImportBatchStatus.COMMITTING
    batch.save(update_fields=["status", "updated_at"])

    valid_rows = list(batch.rows.filter(status=InvoiceImportRowStatus.VALID).order_by("row_number"))
    now = timezone.now()

    created_invoices = 0
    created_vendors = 0
    total_deducted = Decimal("0.00")

    # Cache created vendors during this commit to avoid duplicate creation
    created_vendor_map: Dict[str, Vendor] = {}

    for row in valid_rows:
        v_name = row.vendor_name.strip()
        norm_v = _normalize_text(v_name)

        # 1. Resolve or Auto-Create Vendor
        vendor = row.resolved_vendor
        if not vendor:
            if norm_v in created_vendor_map:
                vendor = created_vendor_map[norm_v]
            else:
                # Check DB once more
                vendor = Vendor.objects.filter(org=batch.org, vendor_name__iexact=v_name).first()
                if not vendor:
                    target_scope = row.resolved_scope_node or ScopeNode.objects.filter(org=batch.org, parent__isnull=True).first()
                    sap_id = f"SAP-AUTO-{slugify(v_name)[:15].upper()}-{uuid4().hex[:6].upper()}"
                    vendor = Vendor.objects.create(
                        org=batch.org,
                        scope_node=target_scope,
                        vendor_name=v_name,
                        sap_vendor_id=sap_id,
                        operational_status=OperationalStatus.ACTIVE,
                        marketing_status=MarketingStatus.APPROVED,
                    )
                    created_vendors += 1
                created_vendor_map[norm_v] = vendor

        # 2. Resolve or Provision Budget Line
        budget = row.resolved_budget
        category = row.resolved_category
        subcategory = row.resolved_subcategory
        budget_line = row.resolved_budget_line

        if not budget_line and budget and category:
            budget_line, _ = BudgetLine.objects.get_or_create(
                budget=budget,
                category=category,
                subcategory=subcategory,
                defaults={
                    "allocated_amount": row.amount or Decimal("0.00"),
                    "consumed_amount": Decimal("0.00"),
                    "reserved_amount": Decimal("0.00"),
                },
            )
            # If line existed with lower allocation than amount, ensure capacity for historical posting
            if budget_line.available_amount < (row.amount or Decimal("0.00")):
                budget_line.allocated_amount += (row.amount or Decimal("0.00"))
                budget_line.save(update_fields=["allocated_amount", "updated_at"])
                budget.allocated_amount += (row.amount or Decimal("0.00"))
                budget.save(update_fields=["allocated_amount", "updated_at"])

        # 3. Generate deterministic invoice number if blank
        inv_num = row.invoice_number.strip()
        if not inv_num:
            inv_num = f"EXP-{slugify(v_name)[:12].upper()}-{row.row_number}"

        target_scope = row.resolved_scope_node or vendor.scope_node
        hist_key = f"bulk:{batch.id}:{row.row_number}:{inv_num}"
        existing_invoice = Invoice.objects.filter(historical_import_key=hist_key).first()
        if not existing_invoice and inv_num and target_scope:
            existing_qs = Invoice.objects.filter(
                scope_node=target_scope,
                vendor=vendor,
                vendor_invoice_number__iexact=inv_num,
            ).exclude(status=InvoiceStatus.REJECTED)
            if category:
                existing_qs = existing_qs.filter(allocations__category=category)
            existing_invoice = existing_qs.first()

        if existing_invoice:
            row.created_invoice = existing_invoice
            row.resolved_vendor = vendor
            row.status = InvoiceImportRowStatus.SKIPPED
            row.error_messages = [
                f"Already posted in system as Invoice #{existing_invoice.id}. Skipped to prevent double deduction."
            ]
            row.save(update_fields=["created_invoice", "resolved_vendor", "status", "error_messages"])
            continue

        title = row.work_description.strip() or f"Invoice {inv_num}"
        amount = row.amount or Decimal("0.00")

        # 4. Create Historical Posted Invoice
        invoice = Invoice.objects.create(
            scope_node=target_scope,
            title=title[:255],
            amount=amount,
            currency="INR",
            status=InvoiceStatus.HISTORICAL_POSTED,
            vendor=vendor,
            created_by=user,
            vendor_invoice_number=inv_num,
            description=row.work_description,
            entry_source=InvoiceEntrySource.HISTORICAL_IMPORT,
            finance_reference_number=f"BULK-{batch.id}-{row.row_number}",
            historical_import_key=f"bulk:{batch.id}:{row.row_number}:{inv_num}",
            historical_posting_reason=f"Bulk historical upload from {batch.file_name}",
            historical_posted_by=user,
            historical_posted_at=now,
        )

        # 5. Create Approved Allocation
        allocation = InvoiceAllocation.objects.create(
            invoice=invoice,
            workflow_instance=None,
            split_step=None,
            entity=target_scope,
            budget=budget,
            category=category,
            subcategory=subcategory,
            amount=amount,
            percentage=Decimal("100.00"),
            status=InvoiceAllocationStatus.APPROVED,
            selected_by=user,
            selected_at=now,
            approved_by=user,
            approved_at=now,
            note=row.work_description,
            metadata={"historical_bulk_import": True, "batch_id": batch.id},
            allocation_source=InvoiceAllocationSource.HISTORICAL_IMPORT,
        )

        # 6. Direct Budget Deduction
        if budget_line:
            source_id = f"invoice:{invoice.id}:allocation:{allocation.id}"
            try:
                res_consume = consume_budget_line_direct(
                    line=budget_line,
                    amount=amount,
                    source_type=SourceType.INVOICE,
                    source_id=source_id,
                    consumed_by=user,
                    note=f"Bulk historical invoice {inv_num}: allocation {allocation.id}",
                )
                allocation.metadata = {
                    "historical_bulk_import": True,
                    "batch_id": batch.id,
                    "budget_line_id": res_consume["budget_line"].id,
                    "budget_consumption_id": res_consume["consumption"].id,
                }
                allocation.save(update_fields=["metadata", "updated_at"])
            except Exception:
                # If direct consumption encounters limit, adjust line and retry
                budget_line.refresh_from_db()
                diff = amount - budget_line.available_amount
                if diff > 0:
                    budget_line.allocated_amount += diff
                    budget_line.save(update_fields=["allocated_amount", "updated_at"])
                    budget.refresh_from_db()
                    budget.allocated_amount += diff
                    budget.save(update_fields=["allocated_amount", "updated_at"])

                consume_budget_line_direct(
                    line=budget_line,
                    amount=amount,
                    source_type=SourceType.INVOICE,
                    source_id=source_id,
                    consumed_by=user,
                    note=f"Bulk historical invoice {inv_num}: allocation {allocation.id}",
                )

        row.created_invoice = invoice
        row.resolved_vendor = vendor
        row.status = InvoiceImportRowStatus.COMMITTED
        row.error_messages = []  # Clear historical error messages
        row.save(update_fields=["created_invoice", "resolved_vendor", "status", "error_messages"])

        created_invoices += 1
        total_deducted += amount

    # Calculate cumulative batch statistics (only count invoices newly posted in this batch)
    committed_rows_qs = batch.rows.filter(
        created_invoice__isnull=False,
        status=InvoiceImportRowStatus.COMMITTED,
    )
    total_committed_invoices = committed_rows_qs.count()
    total_committed_amount = (
        committed_rows_qs.aggregate(total=Sum("amount"))["total"]
        or Decimal("0.00")
    )
    newly_created_vendors_total = (batch.created_vendors_count or 0) + created_vendors

    batch.created_invoices_count = total_committed_invoices
    batch.created_vendors_count = newly_created_vendors_total
    batch.total_amount_deducted = total_committed_amount
    batch.status = InvoiceImportBatchStatus.COMMITTED

    append_batch_activity_log(
        batch,
        event="COMMITTED",
        message=f"Committed batch: {created_invoices} new invoices posted, {created_vendors} new vendors created, ₹{total_deducted:,.2f} deducted. Total posted: {total_committed_invoices}/{batch.total_rows}.",
        user=user,
        extra_details={
            "new_invoices": created_invoices,
            "new_vendors": created_vendors,
            "amount_deducted": str(total_deducted),
            "total_committed": total_committed_invoices,
        },
    )

    batch.save(
        update_fields=[
            "created_invoices_count",
            "created_vendors_count",
            "total_amount_deducted",
            "activity_logs",
            "status",
            "updated_at",
        ]
    )

    return batch


@transaction.atomic
def rerun_invoice_import_batch(batch: InvoiceImportBatch, user=None) -> InvoiceImportBatch:
    """
    Re-evaluates and commits any remaining unposted or previously errored rows in an existing batch.
    - Re-validates the batch (resolves newly added/renamed categories, fuzzy matches, etc.)
    - If valid uncommitted rows are found, commits and deducts them.
    - Idempotently skips already committed rows to ensure 0 duplicate deductions.
    """
    initial_errors = batch.error_rows
    initial_committed = batch.created_invoices_count

    batch = validate_invoice_import_batch(batch)
    if batch.valid_rows > 0:
        batch = commit_invoice_import_batch(batch, user=user)

    resolved_count = max(0, initial_errors - batch.error_rows)
    newly_posted = max(0, batch.created_invoices_count - initial_committed)

    append_batch_activity_log(
        batch,
        event="RERUN",
        message=f"Re-run complete: {resolved_count} previous error(s) resolved, {newly_posted} new invoice(s) posted. Total posted: {batch.created_invoices_count}/{batch.total_rows}.",
        user=user,
        extra_details={
            "errors_resolved": resolved_count,
            "newly_posted": newly_posted,
            "remaining_errors": batch.error_rows,
            "total_posted": batch.created_invoices_count,
        },
    )
    batch.save(update_fields=["activity_logs", "updated_at"])

    return batch

