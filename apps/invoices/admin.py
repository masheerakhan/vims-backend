from django.contrib import admin
from .models import Invoice, InvoiceImportBatch, InvoiceImportRow


@admin.register(Invoice)
class InvoiceAdmin(admin.ModelAdmin):
    list_display = (
        "id", "title", "vendor", "amount", "currency", "status",
        "entry_source", "finance_reference_number", "scope_node", "created_by", "created_at",
    )
    list_filter = ("status", "entry_source", "currency")
    search_fields = ("title", "vendor_invoice_number", "finance_reference_number", "vendor__vendor_name")
    raw_id_fields = (
        "scope_node", "vendor", "created_by", "historical_posted_by", "historical_reversed_by",
    )


class InvoiceImportRowInline(admin.TabularInline):
    model = InvoiceImportRow
    extra = 0
    can_delete = False
    readonly_fields = (
        "row_number",
        "invoice_number",
        "vendor_name",
        "amount",
        "status",
        "resolved_vendor",
        "resolved_scope_node",
        "resolved_budget_line",
        "created_invoice",
    )
    fields = (
        "row_number",
        "invoice_number",
        "vendor_name",
        "amount",
        "status",
        "resolved_vendor",
        "resolved_scope_node",
        "created_invoice",
    )
    show_change_link = True


@admin.register(InvoiceImportBatch)
class InvoiceImportBatchAdmin(admin.ModelAdmin):
    list_display = (
        "id",
        "file_name",
        "org",
        "status",
        "total_rows",
        "valid_rows",
        "error_rows",
        "created_invoices_count",
        "total_amount_deducted",
        "uploaded_by",
        "created_at",
    )
    list_filter = ("status", "org", "created_at")
    search_fields = ("file_name", "id")
    readonly_fields = (
        "created_at",
        "updated_at",
        "total_rows",
        "valid_rows",
        "error_rows",
        "created_invoices_count",
        "created_vendors_count",
        "total_amount_deducted",
        "error_summary",
        "activity_logs",
    )
    raw_id_fields = ("org", "uploaded_by")
    inlines = [InvoiceImportRowInline]


@admin.register(InvoiceImportRow)
class InvoiceImportRowAdmin(admin.ModelAdmin):
    list_display = (
        "id",
        "batch",
        "row_number",
        "invoice_number",
        "vendor_name",
        "amount",
        "status",
        "resolved_vendor",
        "created_invoice",
    )
    list_filter = ("status", "batch__status")
    search_fields = ("invoice_number", "vendor_name", "batch__file_name")
    raw_id_fields = (
        "batch",
        "resolved_vendor",
        "resolved_scope_node",
        "resolved_budget",
        "resolved_category",
        "resolved_subcategory",
        "resolved_budget_line",
        "created_invoice",
    )
    readonly_fields = ("raw_data",)
