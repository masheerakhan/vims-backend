import pytest
from rest_framework import status
from django.core.files.uploadedfile import SimpleUploadedFile

from apps.vendors.models import Vendor, VendorImportBatch, VendorImportBatchStatus
from apps.vendors.import_services import generate_vendor_import_template


@pytest.mark.django_db
class TestVendorBulkImport:
    def test_non_tenant_admin_blocked(self, client_auth, user, org):
        # user has regular role or no tenant_admin
        resp = client_auth.get("/api/v1/vendors/import-batches/")
        assert resp.status_code == status.HTTP_403_FORBIDDEN
        assert "Only Tenant Admin" in str(resp.data)

    def test_template_download(self, client_auth, tenant_admin_user, org):
        client_auth.force_authenticate(user=tenant_admin_user)
        resp = client_auth.get("/api/v1/vendors/import-batches/template/")
        assert resp.status_code == status.HTTP_200_OK
        assert resp.get("Content-Disposition") == 'attachment; filename="Vendor_Import_Template.xlsx"'
        assert len(resp.content) > 1000

    def test_upload_validate_and_commit_batch(self, client_auth, tenant_admin_user, scope_node, org):
        client_auth.force_authenticate(user=tenant_admin_user)
        template_buf = generate_vendor_import_template()
        upload_file = SimpleUploadedFile(
            "test_vendors.xlsx",
            template_buf.getvalue(),
            content_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        
        # Upload
        resp_upload = client_auth.post(
            "/api/v1/vendors/import-batches/upload/",
            {"file": upload_file, "org": org.id},
            format="multipart"
        )
        assert resp_upload.status_code == status.HTTP_201_CREATED
        batch_id = resp_upload.data["id"]
        assert resp_upload.data["status"] in ("validated", "uploaded")
        assert resp_upload.data["total_rows"] >= 1

        # Commit
        resp_commit = client_auth.post(f"/api/v1/vendors/import-batches/{batch_id}/commit/")
        assert resp_commit.status_code == status.HTTP_200_OK
        assert resp_commit.data["status"] == "committed"
        assert resp_commit.data["created_vendors_count"] >= 1
