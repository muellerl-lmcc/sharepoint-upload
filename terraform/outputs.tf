output "tenant_id" {
  description = "Azure Tenant ID"
  value       = data.azuread_client_config.current.tenant_id
}

output "client_id" {
  description = "App Registration Client ID"
  value       = azuread_application.sharepoint_sync.client_id
}

output "client_secret" {
  description = "Client Secret (sensitive)"
  value       = azuread_service_principal_password.sharepoint_sync.value
  sensitive   = true
}
