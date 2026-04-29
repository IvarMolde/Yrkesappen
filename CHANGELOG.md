# Changelog

## v1.1.0 - 2026-04-29

### Added
- API input validation with `zod` for `/api/logginn` and `/api/generer`.
- API rate limiting for general traffic and stricter login protection.
- Health endpoint: `GET /api/health`.
- Centralized server error handler.
- Short-lived signed auth token flow for safer frontend session handling.
- Automated smoke tests (`npm test`) for health, login, and validation.

### Changed
- Refactored backend into smaller modules (routes and services) for maintainability.
- Improved temporary PPTX file handling using OS temp directory + safe cleanup.
- Updated frontend auth flow to use token instead of storing password.
- Fixed frontend text bug where `${valgtNiva}` appeared literally.

### Security
- Reduced exposure of sensitive credentials in browser storage.
- Improved request throttling and validation at API boundaries.
