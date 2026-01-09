#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Tue Jan  6 11:36:28 2026

@author: nr
"""
import argparse
import logging
import os
import re
import shutil as sh
import sys
from nccr_cat_scripts import helpers


# --- Logger Setup ---
logger = logging.getLogger(__name__)
logger.setLevel(logging.INFO) 
handler = logging.StreamHandler()
formatter = logging.Formatter('%(levelname)s:%(message)s')
handler.setFormatter(formatter)
if logger.handlers:
    for h in logger.handlers:
        logger.removeHandler(h)
logger.addHandler(handler)
# --- End Logger Setup ---

class EncodingMismatchError(ValueError):
    """Raised when the detected file encoding does not match the expected one."""
    pass

def decode_scientific(file_path, enc=None):
    # Order of probability for European lab equipment
    encodings = [enc] if enc else ['utf-8', 'latin-1', 'cp1252', 'utf-16']
    
    # REGEX EXPLANATION:
    # \u0000-\u007F: Basic ASCII
    # \u0080-\u00FF: Latin-1 Supplement (µ, °, é, etc.)
    # \u0370-\u03FF: Greek and Coptic characters
    # \s: Whitespace (newlines, tabs)
    valid_pattern = re.compile(r'^[\u0000-\u007F\u0080-\u00FF\u0370-\u03FF\s]*$')

    with open(file_path, 'rb') as f:
        raw_data = f.read()

    for encoding in encodings:
        try:
            # Attempt to decode
            decoded_text = raw_data.decode(encoding)
            
            # Verify if the content fits our "Scientific/European" universe
            if valid_pattern.match(decoded_text):
                logger.debug(f"Successfully decoded this file with {encoding}: {file_path}")
                return encoding, decoded_text
            elif enc:
                raise EncodingMismatchError(f"This file is not encoded with {encoding}: {file_path}")
            else:
                logger.debug(f"Skipping {encoding}: Contains characters outside scientific range.")
        except UnicodeDecodeError:
            logger.debug(f"Could not decode this file with {encoding}: {file_path}")
        except LookupError:
            raise LookupError("Encoding lookup error, maybe a typo?") from None # we use our own message
            
    raise ValueError(f"Could not find a valid encoding that matches the expected character set for {file_path}.")


def process_file(path, enc=None, inplace=None, dest=None, check_dest=True):
    try:
        encoding, decoded_text = decode_scientific(path, enc=enc)
    except EncodingMismatchError as e:
        logger.error(e)
    except ValueError as e:
        logger.error(e)
    except Exception as e:
        logger.error(f"Unknown error encountered for {path}. Error: {e}")
    if inplace:
        dest = path
    elif dest is None:
            base, ext = os.path.splitext(path) # NB ext has a . (e.g. .txt)
            dest = f"{base}_utf8{ext}"
    elif check_dest and helpers.isdir(dest):
        filename = os.path.basename(path)
        base, ext = os.path.splitext(filename) # NB ext has a . (e.g. .txt)
        dest = os.path.join(dest, f"{base}_utf8{ext}")
    if encoding == "utf-8":
        if inplace:
            logger.info(f"File was already in UTF-8: {dest}")
        else:
            sh.copy2(path, dest)
            logger.info(f"File was already in UTF-8, copied from {path} to {dest}")
    else:
        with open(dest, 'w', encoding='utf-8') as f:
            f.write(decoded_text)
        logger.info(f"Converted to UTF-8 {dest}" if inplace else f"Converted {path} into {dest} (UTF-8)")
    
            
def process_recursively(path, formats=None, enc=None, inplace=False, dest=None):
    if dest is None and not inplace:
        dest = f"{path[:-1] if path.endswith(os.sep) else path}_utf8"
        logger.info(f"You neither specified a destination nor used --inplace. Using {dest} as destination")
    if formats is None:
        raise ValueError("")
    for folder, subfolder, files in os.walk(path):
        for file in files:
            fpath = os.path.join(folder, file)
            if inplace:
                outfpath = fpath
            else:
                # Recreate subfolder structure in destination
                rel_path = os.path.relpath(folder, path)
                out_dir = os.path.join(dest, rel_path)
                os.makedirs(out_dir, exist_ok=True)
                outfpath = os.path.join(out_dir, file)
            if file.endswith(formats):
                process_file(fpath, enc=enc,inplace=inplace, dest=outfpath, check_dest=False)
            elif not inplace:
                sh.copy2(fpath, outfpath)
                
                
def run_conversion(args):
    """
    Validation and dispatch logic for the 'convert' command.
    """
    # 1) Check for contradictory arguments
    if args.inplace and args.dest:
        logger.error("Contradictory arguments: You cannot use --inplace and --dest together.")
        sys.exit(1)

    # Convert formats string to tuple
    format_tuple = tuple(f.strip() if f.startswith('.') else f'.{f.strip()}' for f in args.formats.split(','))

    # 2) Dispatch based on path type
    if os.path.isdir(args.path):
        process_recursively(
            path=args.path,
            formats=format_tuple,
            enc=args.enc,
            inplace=args.inplace,
            dest=args.dest
        )
    elif os.path.isfile(args.path):
        process_file(
            path=args.path,
            enc=args.enc,
            inplace=args.inplace,
            dest=args.dest
        )
    else:
        logger.error(f"The path '{args.path}' does not exist.")
        sys.exit(1)

def cli():
    parser = argparse.ArgumentParser(description="File Encoding Converter")
    # Add --log to the main parser so it works for all commands
    parser.add_argument(
        '--log', '--verbosity', '-l', 
        choices=['DEBUG', 'INFO', 'WARNING', 'ERROR', 'CRITICAL'],
        default='INFO',
        help='Set the logging level (default: INFO)'
    )
    # Define a subparser to handle different commands (even if for now it is only one)
    # Use add_subparsers to handle 'process' and 'check' commands
    subparsers = parser.add_subparsers(
        title='commands',
        description='valid commands',
        help='available actions',
        required=True
    )
    parser_convert = subparsers.add_parser(
        'convert', 
        help='Just convert text encoding to UTF-8.'
    )
    parser_convert.set_defaults(func=run_conversion)
    parser_convert.add_argument('path', help="Path to the directory or file to process")
    parser_convert.add_argument('--inplace', action='store_true', help="Overwrite original files")
    parser_convert.add_argument('--dest', type=str, help="Destination path/directory")
    parser_convert.add_argument('--enc', type=str, help="Expected encoding. Use it if you know it, it will make the conversion faster and more robust.")
    parser_convert.add_argument('--formats', type=str, default=".txt", help="Comma-separated extensions")

    args = parser.parse_args()
    numeric_level = getattr(logging, args.log.upper(), logging.INFO)
    logger.setLevel(numeric_level)
    args.func(args)

if __name__ == "__main__":
    cli()
                    