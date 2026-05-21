variable "app_display_name" {
  description = "Anzeigename der Entra ID App Registration"
  type        = string
  default     = "SharePoint-Uploader"
}

variable "secret_expiry_date" {
  description = "Ablaufdatum des Client Secrets (ISO 8601)"
  type        = string
  default     = "2027-01-01T00:00:00Z"
}

variable "write_env_file" {
  description = "Ob die .env Datei automatisch geschrieben werden soll"
  type        = bool
  default     = true
}
