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

# Updating
If you use any functionality from this README and you see it does not work, it might be because you use an old version of the code. To update your code run `git pull`.

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

# Tabular utils
Functionalities to check and correct specific recurring issues with tabular data files. It can handle csv, tsv, xlsx and xls. Nonetheless, formulas within .xls files cannot be read and processed and will be lost if the file is processed. As such, using xlsx is recommended.

The module deals with the following bad practices.
**Padded tables:** tables padded with empty space around them reduce machine readability.
**Trailing spaces:** cells with text with trailing spaces or spaces at the beginning can reduce machine readability.
**Multiple tables:** sheets containing more than one table reduce machine readability. The module uses empty columns and rows to detect and process multiple tables in a sheet. For the moment, only vertically or hortizontally split multiple tables are treated. Presence of both directional splits may be implemented later.

## CLI
The CLI command for tabular utils is `tab-utils`. This must be followed by a command: either `check` or `process`.
Both commands need a positional argument "source" which can be either a file or a folder. In the latter case, it will act on all files within the folder and recursively in its subfolders.

Examples:
```
tab-utils check --strip-unpad file.xlsx
tab-utils check --strip-unpad folder
tab-utils check --multi-table folder
tab-utils process --stip-unpad folder --inplace
tab-utils process --stip-unpad folder --destination folder_processed
tab-utils process --vsplit-tables file.xls --out-format xlsx 
```

### Check options
You will need to select one and only one of these options.
```
--unpad-only  # only check for padded tables
--strip-text  # only check for spaces at the beginning or the end of the cell value
--strip-unpad  # check for both of the above
--multi-table  # check for potential multi-table sheets
```
### Process options
```
--unpad-only  # only unpad the tables. NB: xls would lose their formulas
--strip-text  # only strip the cell value
--strip-unpad  # perform both of the above
--vsplit-tables # split vertically-stacked tables (more details below)
--vsplit-into-two-columns-tables  # split vertically-stacked tables into two columns tables (more details below)
--hsplit-tables  # split horizontally-stacked tables (more details below)
```

You will also need to provide either `--inplace` to edit files in place or `--destination DESTINATION` to provide either a filename (for files) or a folder.
For the processing of multitables you can provide `--out-format [csv|tsv|xlsx|xls]` to specify the output format.

#### unpad
For xlsx, this preserves the cell color and borders, text formatting (e.g. text color, bold, italic) and particularly **formulas** even if across sheets.
Unfortunately, this is not possible for xls.

#### vsplit-tables
This splits the tables at every empty column. If a table-name header is present above the column headers, it will detect the table name and use it to name the sheets. Examples will be provided soon for more clarity.

#### vsplit-into-two-columns-tables
For any block delimited by an empty column, obtains a series of 2-columns tables. e.g. from columns A,B,C,D => [A,B], [A,C], [A,D]. Examples will be provided soon for more clarity.

#### hsplit-tables
This splits the tables at every empty row. If a table-name header is present above the column headers, it will detect the table name and use it to name the sheets. Examples will be provided soon for more clarity.

## Python usage
The main forward-facing functions are listed below.

### Checking
The functions to check are:
```
check_file(file, extension, check_padding, check_strip)  # the last two are booleans about whether you want to check the padding and the stripping
check_recursively(folder, check_padding, check_strip)
check_multitable_file(file, ext)
check_multitable_folder(folder)
```

### Processing
The main functions to process are:
```
unpad_strip_file(file, dest, ext, unpad, strip_text)  # provide dest == file to edit inplace. unpad and strip_text are booleans controlling what you want to perform
unpad_strip_recursively(folder, dest, unpad, strip_text)
vsplit_tables(file, in_format=[extension, optional], out_format=[desired output format, optional], inplace=args.inplace, destination=[destination path, optional])
vsplit_into_two_colum_tables(file, in_format=[extension, optional], out_format=[desired output format, optional], inplace=args.inplace, destination=[destination path, optional])
hsplit_tables(file, in_format=[extension, optional], out_format=[desired output format, optional], inplace=args.inplace, destination=[destination path, optional])
process_recursively(folder,split_func, out_format=None, destination=None, inplace=False)  # split func is one of the 3 functions above, the other arguments are described in the lines above
```


