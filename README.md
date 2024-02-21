# Excel-DirectoryBuilder

## Overview
Excel-DirectoryBuilder is an Excel VBA macro that facilitates the automatic creation of folder structures based on the data specified in an Excel worksheet. This tool streamlines the process of managing large numbers of folders, making it ideal for project setup, file organization, and more.

## Features
- Recursive folder creation based on worksheet hierarchy
- Easy to use with any Excel version that supports VBA
- Customizable base directory and starting cell
- Error handling for non-existent paths

## Getting Started
To get started with Excel-DirectoryBuilder, follow these steps:

1. Download the `.xlsm` file from the releases section.
2. Open the file in Excel (ensure macros are enabled).
3. Adjust the base path in cell C2 to your desired root directory.
4. Enter your folder structure starting from cell B5.
5. Click the 'Create Folders' button to build your directory structure.

## Usage
Simply fill out the hierarchical structure of your desired folders in the Excel sheet and run the macro. The script reads from cell `C2` for the base path and starts creating folders from the structure defined beginning at cell `B5`.

## Contributions
Contributions are welcome! If you'd like to contribute, please fork the repository and use a feature branch. Pull requests are warmly welcome.

## Licensing
This project is released under the MIT License. See the LICENSE file for more details.

## Acknowledgements
- Thanks to the Excel and VBA communities for continuous support and inspiration.
- This project is not affiliated with or endorsed by Microsoft.

## Contact
For bugs, feature requests, or additional questions, please open an issue in the GitHub issue tracker.

Happy folder structuring!
