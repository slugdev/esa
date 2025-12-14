Excel Server Application

A modern web-based platform for hosting and managing Excel-powered applications with a streamlined app store interface.

## Overview

ESA transforms Excel workbooks into interactive web applications with custom user interfaces. Users can browse, launch, and interact with Excel-based apps through an intuitive web interface, while developers can build and deploy their applications with an integrated UI builder.

## Features

### 🎨 Modern App Store Interface
- **Visual App Browser**: Browse applications with prominent images, names, and descriptions
- **Real-time Search**: Instantly filter apps by name, owner, or access group
- **Full-Screen Experience**: Launched apps consume the full viewport for an immersive experience
- **Responsive Design**: Works seamlessly across desktop and mobile devices

### 👤 User Management
- **Role-Based Access**: Support for User, Developer, and Administrator roles
- **Authentication**: Secure login with session management
- **Group-Based Permissions**: Control app access via user groups

### 🛠️ Developer Tools
- **UI Builder**: Visual editor for mapping Excel cells to form inputs and displays
- **App Management**: Create, update, and configure applications
- **Image Upload**: Custom preview images for each application
- **Version Control**: Track application versions

### 📊 Excel Integration
- **Live Workbook Connection**: Real-time data synchronization with Excel
- **Two-Way Data Binding**: Edit inputs to push values back to Excel cells
- **Sheet Navigation**: Support for multi-sheet workbooks
- **Cell Mapping**: Link UI components to specific Excel ranges

## Architecture

### Client (`/client`)
- **Technology**: Vanilla JavaScript, HTML5, CSS3
- **Files**:
  - `index.html` - Main application page
  - `app.js` - Application logic and API integration
  - `style.css` - Modern UI styling
  - `login.html` / `login.js` - Authentication pages
  - `register.html` / `register.js` - User registration

### Server (`/server`)
- **Technology**: C++ with REST API
- **Build System**: CMake
- **Key Files**:
  - `main.cpp` - Server implementation
  - `config.json` - Server configuration
  - `CMakeLists.txt` - Build configuration

## Getting Started

### Prerequisites
- C++ compiler (MSVC, GCC, or Clang)
- CMake 3.15 or higher
- Modern web browser
- Microsoft Excel (for workbook operations)

### Building the Server

```bash
cd server
mkdir build
cd build
cmake ..
cmake --build . --config Release
```

### Running the Server

```bash
cd server/build/Release
./esa
```

The server will start on `http://localhost:8080` by default.

### Accessing the Client

1. Open a web browser
2. Navigate to the client files or serve them via a web server
3. Access `index.html`
4. Login with your credentials or register a new account

## Usage

### For Users
1. **Browse Applications**: View available apps in the Applications tab
2. **Search & Filter**: Use the search bar to find specific apps
3. **Launch Apps**: Click on an app card to open it full-screen
4. **Interact**: Edit input fields to send values to Excel
5. **Refresh Data**: Click "Refresh Data" to pull latest values from Excel
6. **Close**: Click "Close App" or press Escape to exit

### For Developers
1. **Create Application**: Click "Create New" in the Developer tab
2. **Upload Workbook**: Select an .xlsx file and optional preview image
3. **Design UI**: Use the "UI Builder" to map Excel cells to UI components
4. **Configure Components**:
   - Set labels for display names
   - Choose sheet and cell references
   - Select input/display mode
   - Pick input types (text, number)
5. **Save Layout**: Publish your UI design
6. **Test**: Launch the app from the Applications tab

### For Administrators
1. **Manage Users**: View all registered users in the Admin tab
2. **Create Users**: Add new users with specific roles and groups
3. **Assign Roles**: Grant User, Developer, or Administrator privileges
4. **Configure Groups**: Set access groups for fine-grained permissions

## Configuration

Edit `server/config.json` to configure:
- Server port
- Database connection
- Excel integration settings
- CORS policies
- Session timeout

## API Endpoints

### Authentication
- `POST /register` - Create new user account
- `POST /login` - Authenticate user
- `POST /logout` - End session

### Applications
- `GET /apps` - List available applications
- `POST /apps/create` - Create new application
- `POST /apps/update` - Update application
- `GET /apps/ui/get` - Retrieve UI schema

### Excel Operations
- `POST /excel/load` - Open workbook
- `POST /excel/close` - Close workbook
- `POST /excel/query` - Read cell values
- `POST /excel/set` - Write cell values
- `POST /excel/sheets` - List worksheet names

### Users (Admin only)
- `GET /users` - List all users
- `POST /users/create` - Create user
- `POST /users/update` - Update user

## Security

- Session-based authentication with bearer tokens
- Role-based access control (RBAC)
- XSS protection via input escaping
- CORS configuration for API security
- Password hashing (server-side)

## Browser Support

- Chrome/Edge 90+
- Firefox 88+
- Safari 14+
- Opera 76+

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

The MIT License allows you to:
- ✅ Use commercially
- ✅ Modify and distribute
- ✅ Use privately
- ✅ Sublicense

The only requirement is to include the original copyright notice in all copies or substantial portions of the software.

