import io
import re
import csv
from typing import List, Dict, Any, Tuple
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

from django.db import transaction
from django.contrib.auth import get_user_model
from django.core.validators import validate_email
from django.core.exceptions import ValidationError

from apps.core.models import ScopeNode, Organization
from apps.vendors.models import (
    Vendor,
    OperationalStatus,
    MarketingStatus,
    VendorImportBatch,
    VendorImportRow,
    VendorImportBatchStatus,
    VendorImportRowStatus,
)

User = get_user_model()

# Header column definition mapping
COLUMN_MAPPING = {
    "sap_vendor_id": ["sap vendor id", "sap vendor id *", "sap_vendor_id", "sap_code", "sap id", "vendor id"],
    "vendor_name": ["vendor name", "vendor name *", "vendor_name", "company name", "name"],
    "scope_node_code": ["department / scope code", "department / scope code *", "scope_node_code", "scope node code", "department", "scope", "scope code", "department code"],
    "email": ["email", "vendor email", "email address", "contact email"],
    "phone": ["phone", "phone number", "contact phone", "mobile"],
    "pan": ["pan", "pan number", "pan_no"],
    "gstin": ["gstin", "gst number", "gstin_no", "gst"],
    "gst_registered": ["gst registered", "gst registered (yes/no)", "gst_registered"],
    "address_line1": ["address line 1", "address_line1", "address", "street"],
    "city": ["city"],
    "state": ["state"],
    "pincode": ["pincode", "pin code", "zip", "postal code"],
    "country": ["country"],
    "bank_name": ["bank name", "bank_name", "bank"],
    "account_number": ["account number", "account_number", "bank account number", "account no"],
    "ifsc": ["ifsc", "ifsc code", "bank ifsc"],
    "beneficiary_name": ["beneficiary name", "beneficiary_name", "account holder name"],
    "preferred_payment_mode": ["payment mode", "preferred payment mode", "preferred_payment_mode", "payment_mode"],
    "po_mandate_enabled": ["po mandate required", "po mandate required (yes/no)", "po_mandate_enabled", "po mandate"],
    "operational_status": ["operational status", "status", "operational_status"],
}


def generate_vendor_import_template() -> io.BytesIO:
    """
    Generate a styled Excel workbook template with sample data and column instructions.
    """
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Vendor Import"

    # Ensure gridlines are visible
    ws.views.sheetView[0].showGridLines = True

    headers = [
        ("SAP Vendor ID *", 18),
        ("Vendor Name *", 32),
        ("Department / Scope Code *", 26),
        ("Email", 26),
        ("Phone", 18),
        ("PAN", 16),
        ("GSTIN", 20),
        ("GST Registered (Yes/No)", 22),
        ("Address Line 1", 30),
        ("City", 18),
        ("State", 18),
        ("Pincode", 14),
        ("Country", 14),
        ("Bank Name", 24),
        ("Account Number", 22),
        ("IFSC Code", 16),
        ("Beneficiary Name", 26),
        ("Payment Mode", 18),
        ("PO Mandate Required (Yes/No)", 26),
        ("Operational Status", 20),
    ]

    header_font = Font(name="Segoe UI", size=11, bold=True, color="FFFFFF")
    header_fill = PatternFill(start_color="EA580C", end_color="EA580C", fill_type="solid")
    header_align = Alignment(horizontal="center", vertical="center", wrap_text=True)

    thin_border = Border(
        left=Side(style="thin", color="E2E8F0"),
        right=Side(style="thin", color="E2E8F0"),
        top=Side(style="thin", color="E2E8F0"),
        bottom=Side(style="thin", color="E2E8F0"),
    )

    ws.row_dimensions[1].height = 32

    for col_idx, (header_text, col_width) in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col_idx, value=header_text)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = header_align
        cell.border = thin_border
        col_letter = get_column_letter(col_idx)
        ws.column_dimensions[col_letter].width = col_width

    # Sample rows
    sample_rows = [
        [
            "SAP-100234",
            "Acme Industrial Solutions Pvt Ltd",
            "marketing",
            "billing@acmeindustrial.com",
            "+91 9876543210",
            "ABCDE1234F",
            "27ABCDE1234F1Z5",
            "Yes",
            "Plot 42, MIDC Industrial Area",
            "Pune",
            "Maharashtra",
            "411019",
            "India",
            "HDFC Bank",
            "50200012345678",
            "HDFC0001234",
            "Acme Industrial Solutions Pvt Ltd",
            "NEFT/RTGS",
            "No",
            "Active",
        ],
        [
            "SAP-100235",
            "Bright Star Logistics LLP",
            "north",
            "contact@brightstarlogistics.in",
            "+91 9123456780",
            "AABCB5678G",
            "07AABCB5678G1ZP",
            "Yes",
            "Sector 18, Commercial Hub",
            "Gurgaon",
            "Haryana",
            "122002",
            "India",
            "ICICI Bank",
            "001105009876",
            "ICIC0000011",
            "Bright Star Logistics LLP",
            "NEFT/RTGS",
            "Yes",
            "Active",
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


def _normalize_header(header: str) -> str:
    if not header:
        return ""
    h = str(header).strip().lower()
    h = re.sub(r"[\*\_]", " ", h)
    return re.sub(r"\s+", " ", h).strip()


def parse_vendor_import_file(uploaded_file) -> List[Dict[str, Any]]:
    """
    Parse an uploaded .xlsx, .xls, or .csv file into a list of row dictionaries.
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
            norm_h = _normalize_header(raw_h)
            for field, aliases in COLUMN_MAPPING.items():
                if norm_h in aliases or any(_normalize_header(a) == norm_h for a in aliases):
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
            norm_h = _normalize_header(raw_h)
            for field, aliases in COLUMN_MAPPING.items():
                if norm_h in aliases or any(_normalize_header(a) == norm_h for a in aliases):
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


@transaction.atomic
def create_vendor_import_batch(
    org: Organization,
    file_name: str,
    parsed_rows: List[Dict[str, Any]],
    user=None,
) -> VendorImportBatch:
    """
    Create a VendorImportBatch and bulk-insert its child VendorImportRow records.
    """
    batch = VendorImportBatch.objects.create(
        org=org,
        file_name=file_name,
        uploaded_by=user,
        status=VendorImportBatchStatus.UPLOADED,
        total_rows=len(parsed_rows),
    )

    rows_to_create = []
    for r in parsed_rows:
        row_num = r.get("_row_number", len(rows_to_create) + 2)
        
        # Parse boolean flags
        gst_reg_str = str(r.get("gst_registered", "")).strip().lower()
        gst_reg = True if gst_reg_str in ("yes", "true", "1", "y") else (False if gst_reg_str in ("no", "false", "0", "n") else None)

        po_man_str = str(r.get("po_mandate_enabled", "")).strip().lower()
        po_man = True if po_man_str in ("yes", "true", "1", "y") else False

        row_obj = VendorImportRow(
            batch=batch,
            row_number=row_num,
            raw_data=r,
            sap_vendor_id=r.get("sap_vendor_id", ""),
            vendor_name=r.get("vendor_name", ""),
            scope_node_code=r.get("scope_node_code", ""),
            email=r.get("email", ""),
            phone=r.get("phone", ""),
            pan=r.get("pan", "").upper(),
            gstin=r.get("gstin", "").upper(),
            gst_registered=gst_reg,
            address_line1=r.get("address_line1", ""),
            city=r.get("city", ""),
            state=r.get("state", ""),
            country=r.get("country", "") or "India",
            pincode=r.get("pincode", ""),
            bank_name=r.get("bank_name", ""),
            account_number=r.get("account_number", ""),
            ifsc=r.get("ifsc", "").upper(),
            beneficiary_name=r.get("beneficiary_name", ""),
            preferred_payment_mode=r.get("preferred_payment_mode", ""),
            po_mandate_enabled=po_man,
            status=VendorImportRowStatus.PENDING,
            error_messages=[],
        )
        rows_to_create.append(row_obj)

    VendorImportRow.objects.bulk_create(rows_to_create)
    return batch


@transaction.atomic
def validate_vendor_import_batch(batch: VendorImportBatch) -> VendorImportBatch:
    """
    Validate all rows in a VendorImportBatch. Checks mandatory fields, ScopeNode
    resolution, format validity, and intra-batch duplicates.
    """
    batch.status = VendorImportBatchStatus.VALIDATING
    batch.save(update_fields=["status", "updated_at"])

    rows = list(batch.rows.all().order_by("row_number"))
    
    # Preload scope nodes for the org
    scope_nodes_by_code = {
        sn.code.lower(): sn for sn in ScopeNode.objects.filter(org=batch.org)
    }
    # Also index by name
    scope_nodes_by_name = {
        sn.name.lower(): sn for sn in ScopeNode.objects.filter(org=batch.org)
    }

    seen_sap_ids = set()
    valid_count = 0
    error_count = 0
    batch_errors = []

    for row in rows:
        errors = []

        # 1. SAP Vendor ID
        sap_id = row.sap_vendor_id.strip()
        if not sap_id:
            errors.append("SAP Vendor ID is required.")
        else:
            sap_upper = sap_id.upper()
            if sap_upper in seen_sap_ids:
                errors.append(f"Duplicate SAP Vendor ID '{sap_id}' in this upload.")
            seen_sap_ids.add(sap_upper)

        # 2. Vendor Name
        if not row.vendor_name.strip():
            errors.append("Vendor Name is required.")

        # 3. Scope Node / Department Code
        scope_code = row.scope_node_code.strip().lower()
        if not scope_code:
            errors.append("Department / Scope Code is required.")
        else:
            matched_node = scope_nodes_by_code.get(scope_code) or scope_nodes_by_name.get(scope_code)
            if not matched_node:
                errors.append(f"Department / Scope '{row.scope_node_code}' not found in organization.")

        # 4. Email validation
        if row.email.strip():
            try:
                validate_email(row.email.strip())
            except ValidationError:
                errors.append(f"Invalid email address '{row.email}'.")

        # 5. PAN format (optional, 10 alphanumeric if provided)
        if row.pan.strip():
            pan_clean = row.pan.strip().upper()
            if not re.match(r"^[A-Z]{5}[0-9]{4}[A-Z]{1}$", pan_clean):
                errors.append(f"Invalid PAN format '{row.pan}' (expected 10 characters e.g. ABCDE1234F).")

        # 6. GSTIN format (optional, 15 chars)
        if row.gstin.strip():
            gst_clean = row.gstin.strip().upper()
            if not re.match(r"^[0-9]{2}[A-Z]{5}[0-9]{4}[A-Z]{1}[1-9A-Z]{1}Z[0-9A-Z]{1}$", gst_clean):
                errors.append(f"Invalid GSTIN format '{row.gstin}' (expected 15 alphanumeric characters).")

        # 7. IFSC Code format (optional, 11 chars)
        if row.ifsc.strip():
            ifsc_clean = row.ifsc.strip().upper()
            if not re.match(r"^[A-Z]{4}0[A-Z0-9]{6}$", ifsc_clean):
                errors.append(f"Invalid IFSC code '{row.ifsc}' (expected 11 characters e.g. HDFC0001234).")

        if errors:
            row.status = VendorImportRowStatus.ERROR
            row.error_messages = errors
            error_count += 1
            batch_errors.append(f"Row {row.row_number}: {'; '.join(errors)}")
        else:
            row.status = VendorImportRowStatus.VALID
            row.error_messages = []
            valid_count += 1

    VendorImportRow.objects.bulk_update(rows, ["status", "error_messages"])

    batch.valid_rows = valid_count
    batch.error_rows = error_count
    batch.error_summary = batch_errors
    batch.status = VendorImportBatchStatus.VALIDATED
    batch.save(update_fields=["valid_rows", "error_rows", "error_summary", "status", "updated_at"])

    return batch


@transaction.atomic
def commit_vendor_import_batch(batch: VendorImportBatch, user=None) -> VendorImportBatch:
    """
    Commit all valid rows in the batch into the live Vendor Directory.
    Creates new vendors or updates existing ones matching sap_vendor_id.
    """
    if batch.status != VendorImportBatchStatus.VALIDATED:
        batch = validate_vendor_import_batch(batch)

    batch.status = VendorImportBatchStatus.COMMITTING
    batch.save(update_fields=["status", "updated_at"])

    scope_nodes_by_code = {
        sn.code.lower(): sn for sn in ScopeNode.objects.filter(org=batch.org)
    }
    scope_nodes_by_name = {
        sn.name.lower(): sn for sn in ScopeNode.objects.filter(org=batch.org)
    }

    valid_rows = batch.rows.filter(status=VendorImportRowStatus.VALID).order_by("row_number")

    created_count = 0
    updated_count = 0
    committed_rows = []

    for row in valid_rows:
        scope_code = row.scope_node_code.strip().lower()
        scope_node = scope_nodes_by_code.get(scope_code) or scope_nodes_by_name.get(scope_code)
        if not scope_node:
            # Fallback to org root node
            scope_node = ScopeNode.objects.filter(org=batch.org, parent__isnull=True).first()

        sap_id = row.sap_vendor_id.strip()

        # Check existing vendor
        existing_vendor = Vendor.objects.filter(org=batch.org, sap_vendor_id__iexact=sap_id).first()

        op_status_str = str(row.raw_data.get("operational_status", "active")).strip().lower()
        op_status = OperationalStatus.INACTIVE if op_status_str in ("inactive", "suspended", "disabled") else OperationalStatus.ACTIVE

        if existing_vendor:
            # Update live fields
            existing_vendor.vendor_name = row.vendor_name.strip()
            if row.email.strip():
                existing_vendor.email = row.email.strip()
            if row.phone.strip():
                existing_vendor.phone = row.phone.strip()
            if row.pan.strip():
                existing_vendor.pan = row.pan.strip().upper()
            if row.gstin.strip():
                existing_vendor.gstin = row.gstin.strip().upper()
            if row.gst_registered is not None:
                existing_vendor.gst_registered = row.gst_registered
            if row.address_line1.strip():
                existing_vendor.address_line1 = row.address_line1.strip()
            if row.city.strip():
                existing_vendor.city = row.city.strip()
            if row.state.strip():
                existing_vendor.state = row.state.strip()
            if row.pincode.strip():
                existing_vendor.pincode = row.pincode.strip()
            if row.country.strip():
                existing_vendor.country = row.country.strip()
            if row.bank_name.strip():
                existing_vendor.bank_name = row.bank_name.strip()
            if row.account_number.strip():
                existing_vendor.account_number = row.account_number.strip()
                existing_vendor.bank_account_number = row.account_number.strip()
            if row.ifsc.strip():
                existing_vendor.ifsc = row.ifsc.strip().upper()
            if row.beneficiary_name.strip():
                existing_vendor.beneficiary_name = row.beneficiary_name.strip()
            if row.preferred_payment_mode.strip():
                existing_vendor.preferred_payment_mode = row.preferred_payment_mode.strip()
            existing_vendor.po_mandate_enabled = row.po_mandate_enabled
            if scope_node:
                existing_vendor.scope_node = scope_node
            existing_vendor.save()
            updated_count += 1
        else:
            # Create new vendor
            Vendor.objects.create(
                org=batch.org,
                scope_node=scope_node,
                vendor_name=row.vendor_name.strip(),
                sap_vendor_id=sap_id,
                email=row.email.strip(),
                phone=row.phone.strip(),
                pan=row.pan.strip().upper(),
                gstin=row.gstin.strip().upper(),
                gst_registered=row.gst_registered,
                address_line1=row.address_line1.strip(),
                city=row.city.strip(),
                state=row.state.strip(),
                country=row.country.strip() or "India",
                pincode=row.pincode.strip(),
                bank_name=row.bank_name.strip(),
                account_number=row.account_number.strip(),
                bank_account_number=row.account_number.strip(),
                ifsc=row.ifsc.strip().upper(),
                beneficiary_name=row.beneficiary_name.strip() or row.vendor_name.strip(),
                preferred_payment_mode=row.preferred_payment_mode.strip() or "NEFT/RTGS",
                po_mandate_enabled=row.po_mandate_enabled,
                operational_status=op_status,
                marketing_status=MarketingStatus.APPROVED,
            )
            created_count += 1

        row.status = VendorImportRowStatus.COMMITTED
        committed_rows.append(row)

    if committed_rows:
        VendorImportRow.objects.bulk_update(committed_rows, ["status"])

    batch.created_vendors_count = created_count
    batch.updated_vendors_count = updated_count
    batch.status = VendorImportBatchStatus.COMMITTED
    batch.save(update_fields=["created_vendors_count", "updated_vendors_count", "status", "updated_at"])

    return batch
