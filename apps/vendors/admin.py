from django.contrib import admin

from .models import (
    Vendor, VendorInvitation, VendorOnboardingSubmission,
    VendorAttachment, VendorFinanceActionToken, VendorFinanceDecision,
    VendorTrainingVideo, VendorImportBatch, VendorImportRow,
)


@admin.register(VendorInvitation)
class VendorInvitationAdmin(admin.ModelAdmin):
    list_display = ("id", "vendor_email", "vendor_name_hint", "status", "org", "scope_node", "invited_by", "created_at")
    list_filter = ("status",)
    search_fields = ("vendor_email", "vendor_name_hint")
    raw_id_fields = ("org", "scope_node", "invited_by")
    readonly_fields = ("token", "created_at", "updated_at")


@admin.register(VendorOnboardingSubmission)
class VendorOnboardingSubmissionAdmin(admin.ModelAdmin):
    list_display = ("id", "normalized_vendor_name", "normalized_email", "status", "submission_mode", "created_at")
    list_filter = ("status", "submission_mode")
    search_fields = ("normalized_vendor_name", "normalized_email")
    raw_id_fields = ("invitation",)
    readonly_fields = ("created_at", "updated_at")


@admin.register(VendorAttachment)
class VendorAttachmentAdmin(admin.ModelAdmin):
    list_display = ("id", "title", "document_type", "file_name", "submission", "uploaded_by", "created_at")
    raw_id_fields = ("submission", "uploaded_by")


@admin.register(VendorFinanceActionToken)
class VendorFinanceActionTokenAdmin(admin.ModelAdmin):
    list_display = ("id", "submission", "action_type", "expires_at", "used_at", "created_at")
    list_filter = ("action_type",)
    raw_id_fields = ("submission",)
    readonly_fields = ("token", "created_at")


@admin.register(VendorFinanceDecision)
class VendorFinanceDecisionAdmin(admin.ModelAdmin):
    list_display = ("id", "submission", "decision", "sap_vendor_id", "acted_at", "created_at")
    list_filter = ("decision",)
    raw_id_fields = ("submission", "acted_via_token")


@admin.register(Vendor)
class VendorAdmin(admin.ModelAdmin):
    list_display = ("id", "vendor_name", "email", "portal_email", "sap_vendor_id", "operational_status", "marketing_status", "po_mandate_enabled", "org", "created_at")
    list_filter = ("operational_status", "marketing_status", "po_mandate_enabled")
    search_fields = ("vendor_name", "email", "portal_email", "sap_vendor_id")
    raw_id_fields = ("org", "scope_node", "onboarding_submission", "approved_by_marketing")
    readonly_fields = ("created_at", "updated_at")

    def save_model(self, request, obj, form, change):
        super().save_model(request, obj, form, change)
        if change and "email" in form.changed_data:
            from .services import sync_vendor_email_change
            sync_vendor_email_change(obj, obj.email, actor=request.user)


@admin.register(VendorTrainingVideo)
class VendorTrainingVideoAdmin(admin.ModelAdmin):
    list_display = ("id", "title", "is_active", "video_file", "created_by", "created_at", "updated_at")
    list_filter = ("is_active", "created_at")
    search_fields = ("title", "description")
    raw_id_fields = ("created_by",)
    readonly_fields = ("created_at", "updated_at")

    def save_model(self, request, obj, form, change):
        if not obj.created_by_id:
            obj.created_by = request.user
        super().save_model(request, obj, form, change)


class VendorImportRowInline(admin.TabularInline):
    model = VendorImportRow
    extra = 0
    can_delete = False
    readonly_fields = ("row_number", "sap_vendor_id", "vendor_name", "scope_node_code", "email", "status")
    fields = ("row_number", "sap_vendor_id", "vendor_name", "scope_node_code", "email", "status")
    show_change_link = True


@admin.register(VendorImportBatch)
class VendorImportBatchAdmin(admin.ModelAdmin):
    list_display = (
        "id",
        "file_name",
        "org",
        "status",
        "total_rows",
        "valid_rows",
        "error_rows",
        "created_vendors_count",
        "updated_vendors_count",
        "uploaded_by",
        "created_at",
    )
    list_filter = ("status", "org", "created_at")
    search_fields = ("file_name", "id")
    raw_id_fields = ("org", "uploaded_by")
    readonly_fields = (
        "created_at",
        "updated_at",
        "total_rows",
        "valid_rows",
        "error_rows",
        "created_vendors_count",
        "updated_vendors_count",
        "error_summary",
    )
    inlines = [VendorImportRowInline]


@admin.register(VendorImportRow)
class VendorImportRowAdmin(admin.ModelAdmin):
    list_display = (
        "id",
        "batch",
        "row_number",
        "sap_vendor_id",
        "vendor_name",
        "scope_node_code",
        "email",
        "status",
    )
    list_filter = ("status", "batch__status")
    search_fields = ("sap_vendor_id", "vendor_name", "email", "batch__file_name")
    raw_id_fields = ("batch",)
    readonly_fields = ("raw_data",)
