terraform {
  required_providers {
    azuread = {
      source  = "hashicorp/azuread"
      version = "~> 2.50"
    }
  }
}

provider "azuread" {}

# Aktuellen Tenant abrufen
data "azuread_client_config" "current" {}

# Microsoft Graph Service Principal (universell, gleiche App ID in allen Tenants)
data "azuread_application_published_app_ids" "well_known" {}

data "azuread_service_principal" "msgraph" {
  client_id = data.azuread_application_published_app_ids.well_known.result["MicrosoftGraph"]
}

# App Registration
resource "azuread_application" "sharepoint_sync" {
  display_name     = var.app_display_name
  sign_in_audience = "AzureADMyOrg"

  required_resource_access {
    resource_app_id = data.azuread_service_principal.msgraph.client_id

    resource_access {
      id   = data.azuread_service_principal.msgraph.app_role_ids["Sites.ReadWrite.All"]
      type = "Role"
    }
  }
}

# Service Principal
resource "azuread_service_principal" "sharepoint_sync" {
  client_id = azuread_application.sharepoint_sync.client_id
}

# Client Secret
resource "azuread_service_principal_password" "sharepoint_sync" {
  service_principal_id = azuread_service_principal.sharepoint_sync.id
  end_date             = var.secret_expiry_date
}

# Admin Consent: Sites.ReadWrite.All
resource "azuread_app_role_assignment" "sites_readwrite" {
  app_role_id         = data.azuread_service_principal.msgraph.app_role_ids["Sites.ReadWrite.All"]
  principal_object_id = azuread_service_principal.sharepoint_sync.object_id
  resource_object_id  = data.azuread_service_principal.msgraph.object_id
}

# .env Datei lokal schreiben
resource "local_sensitive_file" "env_file" {
  count    = var.write_env_file ? 1 : 0
  filename = "${path.module}/../.env"
  content  = <<-EOT
    AZURE_TENANT_ID=${data.azuread_client_config.current.tenant_id}
    AZURE_CLIENT_ID=${azuread_application.sharepoint_sync.client_id}
    AZURE_CLIENT_SECRET=${azuread_service_principal_password.sharepoint_sync.value}
  EOT
}
