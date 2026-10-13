import os
import sys
import django
from datetime import datetime

# Setup Django environment
os.environ.setdefault("DJANGO_SETTINGS_MODULE", "config.settings.local")
django.setup()

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

from apps.vendors.models import (
    Vendor,
    VendorOnboardingSubmission,
    VendorInvitation,
    SubmissionStatus,
    MarketingStatus,
    OperationalStatus,
)


def export_vendors_to_excel(output_filename: str = "vendors_status_report.xlsx"):
    print("Fetching vendor data from database...")

    wb = openpyxl.Workbook()
    
    # ── Sheet 1: All Registered Vendors ─────────────────────────────────────────
    ws_vendors = wb.active
    ws_vendors.title = "Registered Vendors"

    # Styling helpers
    header_fill = PatternFill(start_color="7B3996", end_color="7B3996", fill_type="solid")
    header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    
    status_approved_fill = PatternFill(start_color="D1FAE5", end_color="D1FAE5", fill_type="solid")
    status_approved_font = Font(name="Calibri", size=10, bold=True, color="065F46")
    
    status_pending_fill = PatternFill(start_color="FEF3C7", end_color="FEF3C7", fill_type="solid")
    status_pending_font = Font(name="Calibri", size=10, bold=True, color="92400E")
    
    status_rejected_fill = PatternFill(start_color="FEE2E2", end_color="FEE2E2", fill_type="solid")
    status_rejected_font = Font(name="Calibri", size=10, bold=True, color="991B1B")

    thin_border = Border(
        left=Side(style="thin", color="E5E7EB"),
        right=Side(style="thin", color="E5E7EB"),
        top=Side(style="thin", color="E5E7EB"),
        bottom=Side(style="thin", color="E5E7EB"),
    )

    headers = [
        "Vendor ID",
        "Vendor Name",
        "Email",
        "Phone",
        "SAP Vendor ID",
        "Park / Branch Scope",
        "Overall Profile Status",
        "Marketing Approval",
        "Operational Status",
        "Profile Revision Pending?",
        "Onboarding Status",
        "GSTIN",
        "PAN",
    ]

    ws_vendors.append(headers)
    for col_idx in range(1, len(headers) + 1):
        cell = ws_vendors.cell(row=1, column=col_idx)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    ws_vendors.row_dimensions[1].height = 26

    # Fetch vendors
    vendors = Vendor.objects.select_related("scope_node", "onboarding_submission").order_by("vendor_name")
    
    for row_idx, v in enumerate(vendors, start=2):
        # Determine overall friendly profile status
        if v.marketing_status == MarketingStatus.APPROVED and v.operational_status == OperationalStatus.ACTIVE:
            profile_status = "Approved"
        elif v.marketing_status == MarketingStatus.PENDING:
            profile_status = "Pending (Marketing Approval)"
        elif v.operational_status == OperationalStatus.WAITING_MARKETING_APPROVAL:
            profile_status = "Pending (Waiting Marketing Approval)"
        elif v.onboarding_submission and v.onboarding_submission.status in (
            SubmissionStatus.SUBMITTED,
            SubmissionStatus.SENT_TO_FINANCE,
            SubmissionStatus.MARKETING_PENDING,
        ):
            profile_status = f"Pending ({v.onboarding_submission.get_status_display()})"
        elif v.marketing_status == MarketingStatus.REJECTED or v.operational_status == OperationalStatus.INACTIVE:
            profile_status = "Rejected / Inactive"
        else:
            profile_status = v.get_operational_status_display()

        scope_name = v.scope_node.name if v.scope_node else "N/A"
        onboarding_status = v.onboarding_submission.get_status_display() if v.onboarding_submission else "Direct / Historical"
        revision_pending = "Yes (Pending Review)" if v.profile_change_pending else "No"

        row_data = [
            v.id,
            v.vendor_name or "—",
            v.email or "—",
            v.phone or "—",
            v.sap_vendor_id or "—",
            scope_name,
            profile_status,
            v.get_marketing_status_display(),
            v.get_operational_status_display(),
            revision_pending,
            onboarding_status,
            v.gstin or "—",
            v.pan or "—",
        ]
        ws_vendors.append(row_data)

        # Style cells
        for col_idx in range(1, len(headers) + 1):
            cell = ws_vendors.cell(row=row_idx, column=col_idx)
            cell.border = thin_border
            cell.font = Font(name="Calibri", size=10)
            cell.alignment = Alignment(vertical="center")

            # Status column coloring (Column 7 = Profile Status)
            if col_idx == 7:
                if "Approved" in profile_status:
                    cell.fill = status_approved_fill
                    cell.font = status_approved_font
                elif "Pending" in profile_status:
                    cell.fill = status_pending_fill
                    cell.font = status_pending_font
                elif "Rejected" in profile_status:
                    cell.fill = status_rejected_fill
                    cell.font = status_rejected_font
                cell.alignment = Alignment(horizontal="center", vertical="center")

            # Marketing Status column coloring (Column 8)
            if col_idx == 8:
                if v.marketing_status == MarketingStatus.APPROVED:
                    cell.fill = status_approved_fill
                    cell.font = status_approved_font
                elif v.marketing_status == MarketingStatus.PENDING:
                    cell.fill = status_pending_fill
                    cell.font = status_pending_font
                elif v.marketing_status == MarketingStatus.REJECTED:
                    cell.fill = status_rejected_fill
                    cell.font = status_rejected_font
                cell.alignment = Alignment(horizontal="center", vertical="center")

        ws_vendors.row_dimensions[row_idx].height = 20

    # Auto-adjust column widths
    for col in ws_vendors.columns:
        max_len = max(len(str(cell.value or "")) for cell in col)
        col_letter = get_column_letter(col[0].column)
        ws_vendors.column_dimensions[col_letter].width = max(max_len + 4, 12)

    # ── Sheet 2: In-Flight Onboarding & Invitations ─────────────────────────────
    ws_invitations = wb.create_sheet(title="In-Flight Onboarding")

    inv_headers = [
        "Invitation ID",
        "Vendor Name Hint",
        "Vendor Email",
        "Scope / Park",
        "Invitation Status",
        "Submission Status",
        "Created At",
        "Invited By",
    ]
    ws_invitations.append(inv_headers)
    for col_idx in range(1, len(inv_headers) + 1):
        cell = ws_invitations.cell(row=1, column=col_idx)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")

    ws_invitations.row_dimensions[1].height = 26

    invitations = VendorInvitation.objects.select_related("scope_node", "invited_by").prefetch_related("submissions").order_by("-created_at")
    for row_idx, inv in enumerate(invitations, start=2):
        latest_sub = inv.submissions.order_by("-id").first()
        sub_status = latest_sub.get_status_display() if latest_sub else "Not Started"
        invited_by_name = inv.invited_by.get_full_name() or inv.invited_by.email if inv.invited_by else "—"

        row_data = [
            inv.id,
            inv.vendor_name_hint or (latest_sub.normalized_vendor_name if latest_sub else "—"),
            inv.vendor_email or "—",
            inv.scope_node.name if inv.scope_node else "—",
            inv.get_status_display(),
            sub_status,
            inv.created_at.strftime("%Y-%m-%d %H:%M") if inv.created_at else "—",
            invited_by_name,
        ]
        ws_invitations.append(row_data)

        for col_idx in range(1, len(inv_headers) + 1):
            cell = ws_invitations.cell(row=row_idx, column=col_idx)
            cell.border = thin_border
            cell.font = Font(name="Calibri", size=10)
            cell.alignment = Alignment(vertical="center")

        ws_invitations.row_dimensions[row_idx].height = 20

    for col in ws_invitations.columns:
        max_len = max(len(str(cell.value or "")) for cell in col)
        col_letter = get_column_letter(col[0].column)
        ws_invitations.column_dimensions[col_letter].width = max(max_len + 4, 14)

    # Save workbook
    wb.save(output_filename)
    abs_path = os.path.abspath(output_filename)
    print(f"\n[SUCCESS] Excel report successfully generated at:\n  {abs_path}")
    print(f"Total Registered Vendors exported: {vendors.count()}")
    print(f"Total In-flight Invitations exported: {invitations.count()}")
    return abs_path


if __name__ == "__main__":
    out_file = sys.argv[1] if len(sys.argv) > 1 else "vendors_status_report.xlsx"
    export_vendors_to_excel(out_file)
