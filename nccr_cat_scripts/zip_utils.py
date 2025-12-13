#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Thu Nov 20 14:53:00 2025

@author: nr
"""
import zipfile as zf
import os
import logging
import tempfile
import shutil
from typing import Dict, Optional
import argparse

# --- Logger Setup (Ensures clean output without '__main__') ---
logger = logging.getLogger(__name__)
logger.setLevel(logging.INFO) 
handler = logging.StreamHandler()
# Define a custom format string that excludes %(name)s (the logger name)
formatter = logging.Formatter('%(levelname)s:%(message)s')
handler.setFormatter(formatter)
if logger.handlers:
    for h in logger.handlers:
        logger.removeHandler(h)
logger.addHandler(handler)
# --- End Logger Setup ---

# Files and folders to strictly ignore during zipping/copying process
SYSTEM_FILES_TO_IGNORE = ['.DS_Store', '__MACOSX', "Thumbs.db"]

def _sanitize_member_path(member, extraction_path):
    """
    Prevents ZipSlip vulnerability by ensuring the extracted path 
    does not escape the intended target directory.
    """
    # 1. Join extraction path and member path
    target = os.path.join(extraction_path, member)
    # 2. Resolve to absolute path
    abs_target = os.path.abspath(target)
    # 3. Check if the absolute target path is still within the extraction path
    if not abs_target.startswith(os.path.abspath(extraction_path) + os.sep):
        # We allow target == extraction_path for cases where the zip contains only '.'
        if abs_target != os.path.abspath(extraction_path):
            # Log and raise an error if a path attempts to escape the extraction directory
            logger.error(f"ZipSlip attempt prevented: {member} tried to escape {extraction_path}")
            raise Exception(f"Potential ZipSlip attempt detected for: {member}")
    return abs_target


def is_single_root_folder(zip_fp):
    """
    Checks if the zip file contains only a single top-level folder (excluding __MACOSX 
    entries), and returns True if so, False otherwise.
    """
    try:
        with zf.ZipFile(zip_fp, "r") as f:
            namelist = f.namelist()
            if not namelist:
                return False # Empty zip file

            # Collect all unique top-level directory names
            top_levels = set()
            for name in namelist:
                # Ignore __MACOSX or .DS_Store
                if True in [i in name for i in SYSTEM_FILES_TO_IGNORE]:
                    continue
                # Split 'folder/file.txt' into 'folder' and 'file.txt'
                root_name = name.split(os.sep, 1)[0]
                # Only consider entries that contain subdirectories or are explicit directories
                if os.sep in name or name.endswith(os.sep):
                    top_levels.add(root_name)

            # Check if there is exactly one top-level component (indicating a wrapped archive)
            if len(top_levels) == 1:
                return True
            
            return False
    except Exception as e:
        # Log error if the zip file inspection fails
        logger.error(f"Error inspecting {os.path.basename(zip_fp)}: {e}")
        return False

def extract_recursively_in_folder(folder, remove_zips=False):
    # Loop continuously to handle archives that extract other archives within the same folder
    while True:
        extracted = set()
        # Find all zip files in the current folder
        zip_fps = [os.path.join(folder, i) for i in os.listdir(folder) if i.endswith("zip")]
        
        for zip_fp in zip_fps:
            # Determine extraction path: 
            # Unwrap (to parent folder) if it contains a single root folder.
            # Otherwise, create a container folder (named after the zip file, minus extension).
            if is_single_root_folder(zip_fp):
                 # Unwrap case: Extract contents directly into the current folder
                extraction_path = os.path.split(zip_fp)[0]
                logger.info(f"-> Unwrapping {os.path.basename(zip_fp)} into {os.path.basename(extraction_path)}/")
            else:
                # Container case: Extract into a new folder named after the zip file
                extraction_path = zip_fp[:-4]
                logger.info(f"-> Creating container and extracting {os.path.basename(zip_fp)} to {os.path.basename(extraction_path)}/")

            # Ensure the target directory exists (for the container case)
            os.makedirs(extraction_path, exist_ok=True)
            
            # Perform the extraction
            try:
                with zf.ZipFile(zip_fp, "r") as f:
                    for member in f.namelist():
                        # Exclude __MACOSX entries during extraction
                        if member.startswith("__MACOSX/"):
                            continue
                        
                        # Sanitize the path before extraction to prevent ZipSlip
                        _sanitize_member_path(member, extraction_path)
                        
                        # Extract the member to the determined path
                        f.extract(member, path=extraction_path)
                
                extracted.add(zip_fp)
                if remove_zips:
                    os.remove(zip_fp)
            
            except zf.BadZipFile:
                logger.error(f"Error: {os.path.basename(zip_fp)} is a bad zip file.")
            except Exception as e:
                logger.error(f"Error extracting {os.path.basename(zip_fp)}: {e}")
        
        # Check if new zip files were created during this iteration. 
        # If no new zips were extracted (the set difference is empty), break the loop.
        current_zip_fps = set(os.path.join(folder, i) for i in os.listdir(folder) if i.endswith("zip"))
        if not(current_zip_fps - extracted):
            break
    
    # After iterative extraction is complete, recursively process all subfolders
    subfolders = [
        full_path 
        for i in os.listdir(folder) 
        if os.path.isdir(full_path := os.path.join(folder, i))
    ]
    
    for subfolder in subfolders:
        extract_recursively_in_folder(subfolder, remove_zips=remove_zips)


def extract_recursively_from_file(filepath, remove_zips=False):
    """
    Handles the initial extraction of a single zip file, 
    then calls the recursive folder processing function.
    """
    
    # Use os.path.splitext to robustly get the base name for the extraction directory
    extraction_dir = os.path.splitext(filepath)[0] 
    if not os.path.exists(extraction_dir):
        os.makedirs(extraction_dir)

    logger.info(f"Initial extract: {os.path.basename(filepath)} to {os.path.basename(extraction_dir)}/")
    
    # Perform extraction member-by-member for safety and exclusion
    try:
        with zf.ZipFile(filepath, "r") as f:
            for member in f.namelist():
                if member.startswith("__MACOSX/"):
                    continue
                _sanitize_member_path(member, extraction_dir)
                f.extract(member, path=extraction_dir)
        
        if remove_zips:
            os.remove(filepath)
    except zf.BadZipFile:
        logger.error(f"Error: {os.path.basename(filepath)} is a bad zip file, aborting recursive extraction.")
        return
    except Exception as e:
        logger.error(f"Error extracting {os.path.basename(filepath)}: {e}")
        return

    # Continue recursively in the newly created folder
    extract_recursively_in_folder(extraction_dir, remove_zips=remove_zips)
    

def extract_recursively(path, remove_zips=False):
    """
    Main entry point: normalizes the path and calls the appropriate handler 
    based on whether the path is a file (ending in .zip) or a folder.
    """
    path = os.path.abspath(path)
    if path.lower().endswith(".zip"):
        extract_recursively_from_file(path, remove_zips=remove_zips)
    else:
        extract_recursively_in_folder(path, remove_zips=remove_zips)
        
        
def _make_naked(zip_fp: str, single_root_folder: str) -> bool:
    """
    Rewrites a 'dressed' zip file (containing only a single folder)
    to a 'naked' zip file (containing the folder's contents).
    This uses temporary disk space to perform the rewrite safely, overwriting the original.
    """
    temp_dir = None
    try:
        # Create a temporary directory for extraction and zipping
        # We ensure it's in the same directory as the zip_fp if possible
        temp_dir = os.path.join(os.path.dirname(zip_fp) or '.', f"temp_naked_clean_{os.path.basename(zip_fp)}_data")
        os.makedirs(temp_dir, exist_ok=True)

        # 1. Extract the content into a temporary root folder
        temp_extract_root = os.path.join(temp_dir, 'temp_root_extract')
        with zf.ZipFile(zip_fp, 'r') as zf_in:
            zf_in.extractall(path=temp_extract_root)

        # Path to the actual content folder that was extracted
        source_content_path = os.path.join(temp_extract_root, single_root_folder.strip('/'))
        
        # 2. Create the new (naked) zip file
        new_zip_fp = os.path.join(temp_dir, 'naked_temp.zip')
        
        with zf.ZipFile(new_zip_fp, 'w', zf.ZIP_DEFLATED) as zf_out:
            for root, _, files in os.walk(source_content_path):
                for file in files:
                    full_path = os.path.join(root, file)
                    # Archive name: path relative to the content folder (flattens the structure)
                    arcname = os.path.relpath(full_path, source_content_path)
                    zf_out.write(full_path, arcname)

        # 3. Replace the original zip file
        shutil.copyfile(new_zip_fp, zip_fp)
        logger.info(f"CLEANED: Made '{os.path.basename(zip_fp)}' naked.")
        return True

    except Exception as e:
        logger.error(f"Failed to make '{os.path.basename(zip_fp)}' naked: {e}")
        return False
    finally:
        # Clean up temporary directories
        if temp_dir and os.path.exists(temp_dir):
             shutil.rmtree(temp_dir)


def _rewrite_zip_for_cleaning(zip_fp: str, temp_zip_fp: str, nested_zips_to_replace: Dict[str, str]):
    """
    Reads from zip_fp, writes cleaned contents to temp_zip_fp.
    - Excludes __MACOSX and .DS_Store entries.
    - Replaces nested zips using files provided in nested_zips_to_replace.
    """
    
    with zf.ZipFile(zip_fp, 'r') as zf_in:
        with zf.ZipFile(temp_zip_fp, 'w', zf.ZIP_DEFLATED) as zf_out:
            
            # Store names of members we have already handled (replaced nested zips)
            handled_members = set()

            for member in zf_in.infolist():
                member_name = member.filename
                
                # A. Skip __MACOSX and .DS_Store
                if True in [i in member_name for i in SYSTEM_FILES_TO_IGNORE] :
                    # Logging done in the main function
                    continue
                
                # B. Check for replacement (cleaned nested zip)
                if member_name in nested_zips_to_replace:
                    cleaned_nested_path = nested_zips_to_replace[member_name]
                    # Add the *cleaned* temporary zip file back to the new archive
                    zf_out.write(cleaned_nested_path, member_name)
                    handled_members.add(member_name)
                
                # C. Copy all other files/directories (including nested zips that weren't replaced)
                elif member_name not in handled_members:
                    # Copy the original member stream to the new zip
                    zf_out.writestr(member, zf_in.read(member))


def clean_zip_recursively(zip_fp: str):
    """
    Recursively cleans a single zip file by removing __MACOSX and .DS_Store entries, 
    making it naked, and cleaning any nested zip files within.
    This process is done in-place by rewriting the zip file.
    """
    zip_filename = os.path.basename(zip_fp)
    
    # 1. Check for Dressed structure and clean it first (in-place rewrite)
    try:
        with zf.ZipFile(zip_fp, 'r') as zf_obj_check:
            single_root = is_single_root_folder(zf_obj_check)
            if single_root:
                _make_naked(zip_fp, single_root)
    except zf.BadZipFile:
        logger.error(f"ERROR: Archive is corrupted and cannot be read: {zip_filename}")
        return

    # 2. Iterate and recursively clean nested zips and identify __MACOSX and .DS_Store
    nested_zips_to_replace: Dict[str, str] = {} # {member_name: path_to_cleaned_temp_zip}
    needs_rewrite = False
    
    try:
        with zf.ZipFile(zip_fp, 'r') as zf_in:
            for member in zf_in.infolist():
                
                # Check A: __MACOSX (requires rewrite)
                if True in [i in member.filename for i in SYSTEM_FILES_TO_IGNORE]:
                    logger.info(f"ISSUE: Found __MACOSX or .DS_Store entry in {zip_filename}: {member.filename}")
                    needs_rewrite = True 
                    
                # Check B: Nested Zip
                if member.filename.lower().endswith('.zip'):
                    nested_zip_name = member.filename
                    needs_rewrite = True # Replacement of a nested zip requires rewrite
                    
                    # Create a temporary file to hold the nested zip's stream for cleaning
                    temp_nested_file = None
                    try:
                        temp_nested_file = os.path.join(tempfile.gettempdir(), f"nested_{os.path.basename(nested_zip_name)}_{os.getpid()}")
                        # Ensure we write the bytes of the nested zip to the temporary file
                        with open(temp_nested_file, 'wb') as tmp:
                            tmp.write(zf_in.read(member))
                        
                        # Recursively clean the temporary file (in-place on temp_nested_file)
                        clean_zip_recursively(temp_nested_file)
                        nested_zips_to_replace[nested_zip_name] = temp_nested_file

                    except Exception as e:
                        logger.error(f"Failed to clean nested zip {nested_zip_name} inside {zip_filename}: {e}")
                        # If cleaning fails, we skip replacement and let the uncleaned zip be copied.
                    
    except Exception as e:
        logger.error(f"Unexpected error during recursive scan of {zip_filename}: {e}")
        return
    
    # 3. If any changes were detected (MACOSX found or nested zips cleaned), rewrite the main zip.
    if needs_rewrite:
        temp_cleaned_zip = None
        try:
            # Create a temporary path for the rewritten archive
            temp_cleaned_zip = os.path.join(tempfile.gettempdir(), f"final_clean_{zip_filename}_{os.getpid()}")
            
            # Perform the rewrite (skipping __MACOSX and .DS_Store, replacing nested zips)
            _rewrite_zip_for_cleaning(zip_fp, temp_cleaned_zip, nested_zips_to_replace)
            
            # Replace the original zip file
            shutil.copyfile(temp_cleaned_zip, zip_fp)
            logger.info(f"SUCCESS: Finished cleaning and rewriting {zip_filename}.")
            
        except Exception as e:
            logger.error(f"Failed to finalize rewrite of '{zip_filename}': {e}")
        finally:
            # Clean up temporary files
            if temp_cleaned_zip and os.path.exists(temp_cleaned_zip):
                 os.remove(temp_cleaned_zip)
            # Clean up all temporary nested zip files
            for path in nested_zips_to_replace.values():
                if os.path.exists(path):
                    os.remove(path)


def main_cleaner(filepath: str, output_filepath: Optional[str] = None, in_place: bool = True):
    """
    Main entry point for recursive cleaning. 
    It copies the source zip to a temporary file, cleans it, and then moves
    the result to either the source path (in-place=True) or output_filepath.
    
    Args:
        filepath (str): The path to the source zip file.
        output_filepath (str, optional): The path to save the cleaned file. 
                                         Required if in_place=False.
        in_place (bool): If True, the original file is overwritten. 
                         If False, the cleaned file is saved to output_filepath.
    """
    if not os.path.exists(filepath):
        logger.error(f"File not found: {filepath}")
        return

    filepath = os.path.abspath(filepath)
    filename = os.path.basename(filepath)
    
    if not filename.lower().endswith('.zip'):
        logger.error(f"Input file must be a zip archive (.zip): {filename}")
        return

    if not in_place and not output_filepath:
        logger.error("Must specify 'output_filepath' when 'in_place' is False.")
        return

    logger.info(f"--- Starting recursive cleaning of {filename} (In-Place: {in_place}) ---")
    logger.warning("Cleaning of the zip file can take quite some time, please be patient and do not interrupt the script")
    
    # 1. Create a working copy of the source file in a temporary directory
    work_fp = os.path.join(tempfile.gettempdir(), f"working_copy_{filename}_{os.getpid()}")
    
    try:
        shutil.copyfile(filepath, work_fp)
        
        # 2. Perform the recursive cleaning on the working copy
        clean_zip_recursively(work_fp)
        
        # 3. Determine the final destination
        final_dest = filepath if in_place else output_filepath
        
        # 4. Move the cleaned working copy to the final destination
        if os.path.exists(final_dest):
            os.remove(final_dest) # Remove existing file before renaming/moving
        shutil.move(work_fp, final_dest)
        
        if in_place:
            logger.info(f"--- Cleaning complete: {filename} was cleaned in-place. ---")
        else:
            logger.info(f"--- Cleaning complete: Saved cleaned file to {os.path.basename(final_dest)} ---")
            
    except Exception as e:
        logger.error(f"An error occurred during cleanup of {filename}: {e}")
        
    finally:
        # Clean up the temporary working file if it still exists
        if os.path.exists(work_fp):
            os.remove(work_fp)



def zip_appropriately(source_dir: str, target_dir: str):
    """
    Creates a copy of the source directory structure in the target directory,
    where:
    1. All subfolders are replaced by a "naked" zip file named after the folder.
    2. Any loose files in the source root are copied directly to the target root.
    
    System files like .DS_Store are excluded from both zipping and copying.
    """
    
    if not os.path.isdir(source_dir):
        logger.error(f"Source directory not found: {source_dir}")
        return

    # Create the target directory, overwriting if it exists to ensure a clean slate
    if os.path.exists(target_dir):
        shutil.rmtree(target_dir)
    os.makedirs(target_dir, exist_ok=True)
    
    logger.info(f"--- Starting naked zipping of '{os.path.basename(source_dir)}' to '{os.path.basename(target_dir)}' ---")

    for item_name in os.listdir(source_dir):
        source_path = os.path.join(source_dir, item_name)

        # Skip system files found at the root level (covers both folders and loose files)
        if item_name in SYSTEM_FILES_TO_IGNORE:
            logger.info(f"  Skipping root system entry '{item_name}'")
            continue
        
        if os.path.isdir(source_path):
            # 1. Handle Folders: Create a naked zip file
            zip_fp = os.path.join(target_dir, f"{item_name}.zip")
            logger.info(f"  Zipping folder '{item_name}/' -> '{item_name}.zip' (naked)")
            
            try:
                with zf.ZipFile(zip_fp, 'w', zf.ZIP_DEFLATED) as zf_out:
                    # os.walk is used to iterate recursively inside the folder
                    for root, _, files in os.walk(source_path):
                        for file in files:
                            
                            # Skip system files found inside subfolders
                            if file in SYSTEM_FILES_TO_IGNORE:
                                logger.info(f"  Skipping nested system file '{file}' in {root}")
                                continue
                            
                            full_path = os.path.join(root, file)
                            # Key part for 'naked' zipping: arcname is relative to source_path,
                            # flattening the top-level folder structure inside the zip.
                            arcname = os.path.relpath(full_path, source_path)
                            zf_out.write(full_path, arcname)
            except Exception as e:
                logger.error(f"Failed to create zip for {item_name}: {e}")

        elif os.path.isfile(source_path):
            # 2. Handle Files: Copy free file
            target_path = os.path.join(target_dir, item_name)
            logger.info(f"  Copying file '{item_name}'")
            try:
                shutil.copy2(source_path, target_path) # copy2 preserves metadata
            except Exception as e:
                logger.error(f"Failed to copy file {item_name}: {e}")
        
    logger.info("--- Naked zipping process complete ---")

def cli():
    """Configures and runs the command line interface."""
    parser = argparse.ArgumentParser(
        description="A utility for extracting, compressing, and cleaning zip archives, particularly with relation to Zenodo datasets.",
        # help="""Use one of the following commands: 
        #     clean: process a zip file and remove any system file or redundant folder tree level,
        #     zip: zip any folder at a given path without redundant levels,
        #     extract: extract recursively ignoring system files
            
        # """
    )
    # Define a subparser to handle different commands (clean or naked)
    subparsers = parser.add_subparsers(dest='command', required=True)

    # --- 'clean' command parser ---
    parser_clean = subparsers.add_parser(
        'clean',
        help='Recursively cleans a zip file, removing system files and redundant single-root wrapping folder.',
        formatter_class=argparse.RawTextHelpFormatter
    )
    parser_clean.add_argument(
        'filepath',
        type=str,
        help='Path to the zip file to be cleaned.'
    )
    parser_clean.add_argument(
        '--output-filepath',
        type=str,
        default=None,
        help='If provided, the cleaned file is saved to this path.'
    )
    parser_clean.add_argument(
        '--in-place',
        action='store_true',
        default=False,
        help="""Clean the file in place, i.e. overwrite the zip file."""
    )
    # Set the function to call when 'clean' is used
    parser_clean.set_defaults(func=handle_clean_command)

    # --- 'zip' command parser ---
    parser_zip = subparsers.add_parser(
        'zip',
        help='Creates a directory of "naked" (i.e. no redundant folder level) zip files from a source directory structure.'
    )
    parser_zip.add_argument(
        'source_dir',
        type=str,
        help='Path to the source directory containing files and folders to zip.'
    )
    parser_zip.add_argument(
        'target_dir',
        type=str,
        help='Path to the target directory where the resulting files and zips will be saved.'
    )
    # --- 'extract' command parser ---
    parser_extract = subparsers.add_parser(
        'extract',
        help='Recursively extracts nested zip files from a file or directory.'
    )
    parser_extract.add_argument(
        'path',
        type=str,
        help='Path to the zip file OR directory to start recursive extraction.'
    )
    parser_extract.add_argument(
        '--remove-zips',
        action='store_true',
        help='Deletes the source zip files after successful extraction.'
    )
    parser_extract.set_defaults(func=handle_extract_command)
    # Set the function to call when 'naked' is used
    parser_zip.set_defaults(func=handle_naked_command)
    
    # Parse the arguments and call the handler function
    args = parser.parse_args()
    args.func(args)


def handle_clean_command(args):
    """Handler function for the 'clean' command."""
    output_filepath = args.output_filepath
    in_place = args.in_place
    if not in_place and not output_filepath:
        logger.error("Error: either specify '--output-filepath' or use '--in-place'.")
        return
        
    if output_filepath and in_place:
        # If both are specified, prioritize the explicit output path
        logger.info("You specified an output path but also '--in-place', the output path will be used.")
        in_place = False
        
    main_cleaner(args.filepath, output_filepath=output_filepath, in_place=in_place)


def handle_naked_command(args):
    """Handler function for the 'naked' command."""
    zip_appropriately(args.source_dir, args.target_dir)

def handle_extract_command(args):
    """Handler function for the 'extract' command."""
    extract_recursively(args.path, args.remove_zips)


if __name__ == '__main__':
    cli()
