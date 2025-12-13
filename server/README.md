# ESA (Excel Server Application)

A minimal REST-style Windows server that manages Excel workbooks via COM. ESA hosts uploaded Excel apps with versioning, access control, and basic Excel automation for querying and updating cell ranges.

## Features
- User login/logout with bearer tokens stored in-memory.
- App CRUD with owners, public/group access control, and versioned .xlsx storage.
- Admin-controlled user management (roles: user, developer, admin from config).
- Excel pool: load workbook per session token, query ranges, set cell values, close/restart Excel instance.
- Naive JSON parsing and raw HTTP over Winsock (demo-grade; trusted environments only).

## Build
```powershell
cd server\build
cmake .. -G "Visual Studio 17 2022" -A x64
cmake --build . --config Debug
```
The binary outputs as `esa.exe`.

## Configuration
`config.json` example:
```json
{
  "port": 8080,
  "excel_instances": 2,
  "users": [{"username": "admin", "password": "admin"}],
  "admins": ["admin"]
}
```
- `admins` users are forced to Admin role even if edited elsewhere.

## API Overview
Headers: `Authorization: Bearer <token>` for authenticated routes. Content-Type `application/json` required for POST/PUT bodies.

Auth
- POST `/login` {"username","password"}
- POST `/logout`

Apps
- GET `/apps`
- POST `/apps` {name, description, file_base64, public?, access_group?}
- PUT `/apps/{name}` {new_version?, description?, file_base64?, public?, access_group?}
- DELETE `/apps/{name}`

Users (admin only)
- GET `/users`
- POST `/users` {username, password, groups?, role?}

Excel
- POST `/excel/load` {owner?, name, version?}
- POST `/excel/query` {sheet, range}
- POST `/excel/set` {sheet, range, value | value_number | value_bool}
- POST `/excel/close`

## Storage Layout
- `app/<owner>/<app>/<version>/` stores uploaded `.xlsx` and `meta.txt`.
- `db.bin` stores users/apps in a simple binary format.

## Notes & Warnings
- No TLS, no rate limiting, naive JSON parsing; use behind trusted network or proxy.
- Excel automation uses COM late binding; requires Excel installed and registered.
- Closing a session restarts the Excel instance to keep the pool clean.
