# useful_scripts

This repository groups useful scripts for the general NCCR Catalysis community.
You can keep up to date on addition to this repo with the "Watch" functionality on GitHub.

If you are familiar with Python, you can use the functions in these scripts as you wish.
If you are not familiar with Python, just make sure to have a Python3 installation and use the scripts with the command line interface (CLI) as explained in this documentation.
The CLI

# Installation

```
# 1. Clone your repo
git clone https://github.com/YourUsername/useful_scripts.git
cd useful_scripts

# 2. Install the package in "editable" mode for development (best practice)
pip install -e .
```

# zip_utils

Functionalities concerning zip files particularly with regards to Zenodo datasets.

## Recursive extraction

The function `extract_recursively(path)` takes a filepath and extracts everything that can be extracted, recursively. It ignores system files such as "\_\_MACOSX" and ".DS_Store".

If the filepath is a folder, it will look for zip files anywhere in that folder or subfolders. If the zip files contains zip files, those will be extracted too.

If the filepath is a .zip file (e.g. a Zenodo dataset downloaded), it will extract its content and look for zip files within it.

To use the CLI:
`zip-utils extract --path /path/to/extract [--remove-zips]`
where `/path/to/extract` is the filepath to extract and the optional flag `--remove-zips` removes the zip files after successful extraction.

## Zipping

Use of this function is recommended to upload to Zenodo: just run it on your dataset folder, open the output, and drag everything onto Zenodo. This avoids system files and redundant folder levels.
The function `zip_appropriately(input_path, output_path)` copies files from `input_path` to `output_path` and zips subfolders and places them in `output_path`.
To use the CLI:
`zip-utils zip --source_dir /path/to/source/dir --target_dir /path/to/target/dir`

## Cleaning

Generally speaking, if using the function `zip_appropriately` to zip, there should be no need to clean up archives, but otherwise the `main_cleaner` function can be used to remove any redundant level and any system file.

To use the function:
use either `main_cleaner(input_filepath, output_filepath=output_filepath)` to obtain the cleaned zip file at `output_filepath` or `main_cleaner(input_filepath, in_place=in_place)` to overwrite the input file.

For the CLI:
`zip-utils clean /path/to/file --output-filepath /path/to/output` or `zip-utils clean /path/to/file --in-place`
