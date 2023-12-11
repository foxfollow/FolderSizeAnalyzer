# FolderSizeAnalyzer
Used for analyze folder and subfolder for their size, than make excel file as output. For Windows

## Usage
Main .exe file in `dist\`

Use `.exe` in folder `dist` or `.py` (Note: for using `.py` should install requirments) 

### Run script 1 exmaple 
```
searchFoldersAnalyzer.exe C:\ Windows "OneDrive Personal"
```

where `C:\` folder to search and `Windows` and `OneDrive Personal` is folders to skip

### Run script 2 exmaple
```
searchFoldersAnalyzer.exe
```
scipt will ask which folder to scan and which folders to skip

## Building
- Check Requirments for building
- Run `build.ps1`
## Requirments
Windows
Tested on Python 3.10
Python libraries `openpyxl` 

### Installing requirments
```
pip install openpyxl
```

### Requirments for building
`pyinstaller` and `openpyxl` 

ensure that you have correct path to your pyinstaller and change `build.ps1` first row if needed


## MIT License

Copyright (c) 2023 Heorhii Savoiskyi