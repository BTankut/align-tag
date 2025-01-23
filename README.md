# Align Tag

A Revit add-in that intelligently aligns annotation tags.

## Features

- **Align Left**: Aligns tags to the left
- **Align Right**: Aligns tags to the right
- **Align Center**: Centers tags (with smart grouping)
- **Distribute**: Distributes tags with equal spacing
- **Arrange**: Automatically arranges tags around the view frame

## Requirements

- Windows 10 or later
- .NET Framework 4.8
- Any Revit version from 2019 to 2024
- For development:
  - Visual Studio 2022 or VS Code
  - C# Dev Kit (for VS Code)
  - .NET SDK 9.0 or later

## Installation

1. Copy these files to your Revit add-ins folder:
   - AlignTag.addin
   - AlignTag.dll

   Path: %APPDATA%\Autodesk\Revit\Addins\[YEAR]
   
   Example: C:\Users\[USERNAME]\AppData\Roaming\Autodesk\Revit\Addins\2022

2. Restart Revit

## Development

### Using Visual Studio 2022:

1. Open the solution (AlignTag.sln)
2. Select desired Revit version from Configuration (e.g., 2022Debug)
3. Build the solution
4. Start debugging

### Using VS Code:

1. Open the project in VS Code
2. Accept using C# Extension
3. Press F5 or use Terminal > Run Build Task to start building
4. After build, the add-in will be automatically copied to the Revit folder

## Supported Revit Versions

- Revit 2019
- Revit 2020
- Revit 2021
- Revit 2022
- Revit 2023
- Revit 2024

## License

MIT License

Copyright (c) 2025 BT
