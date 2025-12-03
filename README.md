# nccr_cat_scripts

This repository groups useful scripts for the general NCCR Catalysis community.
You can keep up to date on addition to this repo with the "Watch" functionality on GitHub.

If you are familiar with Python, you can use the functions in these scripts as you wish.
If you are not familiar with Python, just make sure to have a Python3 installation and use the scripts with the command line interface (CLI) as explained in this documentation.

For any feedback, to report a bug, or to request a feature, contact: [admin@nccr-catalysis.ch](mailto:admin@nccr-catalysis.ch)

1\. Installation Guide for Beginners
------------------------------------

To install and use these scripts, you only need two prerequisites: **Python** and **Git**. You **do not** need a GitHub account, as this repository is public.

### Step 1: Install Python (and Pip)

Python 3 is required. The installation process for Python typically includes **Pip** (the package installer), which we will use to install the scripts.

| **Operating System** | **How to Install Python** |
| --- | --- |
|   **Windows**   |   1\. Go to the [Official Python Website](https://www.python.org/downloads/windows/ "null").      2\. Download the latest Python 3 installer.      3\. **CRUCIAL STEP:** During installation, make sure to check the box that says **"Add Python to PATH"** before clicking "Install Now."   |
|   **macOS**   |   1\. Go to the [Official Python Website](https://www.python.org/downloads/macos/ "null").      2\. Download the latest Python 3 installer.      3\. Run the installer package.   |
|   **Linux (Ubuntu/Debian)**   |   Most distributions come with Python pre-installed. You can ensure you have Python 3 and Pip by running the following commands in your terminal:      `sudo apt update`      `sudo apt install python3 python3-pip`   |

**Verification:** Open your command line (Terminal on Mac/Linux, Command Prompt/PowerShell on Windows) and type:

    python3 --version
    # You should see something like: Python 3.10.12
    
### Step 2: Install Git

Git is required to download (clone) the code from GitHub to your local machine.

| **Operating System** | **How to Install Git** |
| --- | --- |
|   **Windows**   |   1\. Go to the [Git website](https://git-scm.com/download/win "null").      2\. Download and run the installer.      3\. Accept all default settings during installation.   |
|   **macOS**   |   The easiest way is to install Apple's Command Line Tools by running:      `xcode-select --install`      Alternatively, you can install Git via [Homebrew](https://brew.sh/ "null").   |
|   **Linux (Ubuntu/Debian)**   |   Run the following command in your terminal:      `sudo apt install git`   |

**Verification:** Open your command line and type:

    git --version
    # You should see something like: git version 2.34.1
    

### Step 3: Install the Scripts

Once Python and Git are ready, follow these steps to install the `nccr_cat_scripts` package:
1.  **Get a terminal in a folder of your choosing:**
Use one of these possible methods:
    1. Use your file system to navigate to the desired folder and then open a terminal there:
        1. right click on the empty space in the folder or on a folder itself, then for  click on "Open PowerShell window here"/"Open command window here"/"Open Terminal Here"/"Open in Terminal" according to your OS
        2. otherwise on Mac control-click (or right-click) the folder name/icon in the Path Bar at the bottom of the window, then choose "Open in Terminal"
    2. Open a terminal and navigate to the folder using the following commands:
        1. To Change Directory to a folder: Use `cd <folder_name>`.
        2. To Move Up one directory: Use `cd ..`.
        3. To See your Current Location: Use `pwd` (Linux/Mac) or `cd` (Windows).
        4. To List files in the current directory: Use `ls` (Linux/Mac) or `dir` (Windows).

2.  **Clone the Repository (Download the code):** This command downloads all the files from GitHub into a folder called nccr_cat_scripts which is placed within the folder you are in now.
    
        git clone https://github.com/nccr_cat_scripts/nccr_cat_scripts.git
        
    
3.  **Navigate to the Directory:** Change into the newly created project folder.
    
        cd nccr_cat_scripts
        
    
4.  **Install the Package:** Use the primary `pip` command, or the Windows workaround if needed. This installs the module in "editable" mode (`-e`).
    
        pip install -e .
    
### **Troubleshooting: 'pip' command not found (Windows)**

If the installation step "Add Python to PATH" was missed on Windows, the system won't recognize the `pip` command.

Solution 1: Use py -m prefix (Recommended)

Instead of typing pip install -e ., you can tell the Python launcher utility (py) to run the Pip module:

    py -m pip install -e .
    

Solution 2: Manually Add Python to PATH

If Solution 1 fails, you need to manually add the Python installation directories to your system's PATH.

1.  **Locate Folders:** Find the two folders in your Python installation (example paths shown):
    
    *   The main Python folder: `C:\Users\YourUser\AppData\Local\Programs\Python\Python310`
        
    *   The Scripts folder: `C:\Users\YourUser\AppData\Local\Programs\Python\Python310\Scripts`
        
2.  **Open Environment Variables:** Press the **Windows Key** and type `environment variables`. Click **"Edit the system environment variables"**, then click the **"Environment Variables..."** button.
    
3.  **Edit PATH:** Under "System variables," select **`Path`** and click **"Edit..."**.
    
4.  **Add New Paths:** Click **"New"** and add the main Python folder path. Click **"New"** again and add the `Scripts` folder path.
    
5.  **Restart Terminal:** Close and reopen your terminal window for the changes to take effect. You should now be able to run `pip install -e .`.
# Importing
To use the functions in your own scripts, import as follows:

```from nccr_cat_scripts import [module name]```

where `[module name]` is any of the modules documented below (zip_utils, excel_utils).

# Command Line Interface (CLI)
Every module can be called in the command line. Generally this is just the module name but where "_" has been replaced with "-".
Then a command needs to be given to specify what function to run. You can see which commands are available by running 

```module-name --help```

or by reading the section of this documentation that concerns that module.
If you do not know how to use a command you can run

`module-name [command] --help`

or check this documentation.

# zip_utils

Functionalities concerning zip files particularly with regards to Zenodo datasets.

## Recursive extraction

The function `extract_recursively(path)` takes a filepath and extracts everything that can be extracted, recursively. It ignores system files such as "\_\_MACOSX" and ".DS_Store".

If the filepath is a folder, it will look for zip files anywhere in that folder or subfolders. If the zip files contains zip files, those will be extracted too.

If the filepath is a .zip file (e.g. a Zenodo dataset downloaded), it will extract its content and look for zip files within it.

To use the CLI:

```
zip-utils extract --path /path/to/extract [--remove-zips]
```

where `/path/to/extract` is the filepath to extract and the optional flag `--remove-zips` removes the zip files after successful extraction.

## Zipping

Use of this function is recommended to upload to Zenodo: just run it on your dataset folder, open the output, and drag everything onto Zenodo. This avoids system files and redundant folder levels.
The function `zip_appropriately(input_path, output_path)` copies files from `input_path` to `output_path` and zips subfolders and places them in `output_path`.

To use the CLI:

```
zip-utils zip --source_dir /path/to/source/dir --target_dir /path/to/target/dir
```

## Cleaning

Generally speaking, if using the function `zip_appropriately` to zip, there should be no need to clean up archives, but otherwise the `main_cleaner` function can be used to remove any redundant level and any system file.

To use the function:
use either `main_cleaner(input_filepath, output_filepath=output_filepath)` to obtain the cleaned zip file at `output_filepath` or `main_cleaner(input_filepath, in_place=in_place)` to overwrite the input file.

To use the CLI:

`zip-utils clean /path/to/file --output-filepath /path/to/output` or `zip-utils clean /path/to/file --in-place`

# Excel utils
Functionalities to clean up specific recurring issues with .xlsx files. 

## Unpad
Tables padded with empty space around them reduce machine readability. You can process a single .xlsx or all the .xlsx inside a folder and its subfolders recursively and remove any padding. This preserves the cell color and borders, text formatting (e.g. text color, bold, italic) and particularly **formulas** even if across sheets.
To use the CLI:

```
excel-utils unpad --source /path/to/source/dir --target /path/to/target/dir
```

## Strip text
Cells with text with trailing spaces or spaces at the beginning can reduce machine readability. You can process a single .xlsx or all the .xlsx inside a folder and its subfolders recursively and remove any space at the beginning and end of text in a cell.
To use the CLI:
```
excel-utils strip-text --source /path/to/source/dir --target /path/to/target/dir
```

## Clean
This just consists in unpadding and stripping the text at the same time. You can run this function on your data before uploading to Zenodo.

To use the CLI:
```
excel-utils clean --source /path/to/source/dir --target /path/to/target/dir
```

## Unpad, strip, or both in a script
The functions process_folder and process_excel_file handle the unpadding, stripping, or both.
Use:

```process_folder(source_fol: str, dest_fol: str, unpad: bool, strip_text: bool)```

or 

```process_excel_file(filename: str, outname: str, unpad: bool, strip_text: bool)```

where the booleans `unpad` and `strip_text` control which operation(s) to perform.

## Check
You can use the function  check_folder_recursively to know if any file in it has issues of either padding or text that needs stripping.
Use:
``` check_folder_recursively(folder_path: str, check_padding: bool, check_strip: bool)```


To use the CLI:
```
excel-utils unpad folder  [--padding] [--strip]
```

where the two optional flags allow to check only for that type of issue.

