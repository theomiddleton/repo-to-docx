"""
This script uses code-convert.py to generate a single Markdown file from
a repository, then uses md_to_word.py to convert it to a syntax-highlighted Word doc.
"""
import os
import sys
import json
from code_convert import create_markdown_from_repo
from md_to_word import convert_markdown_to_docx

CONFIG_FILE = "code_convert_config.json"

def load_config():
    """Load configuration from a local JSON file if it exists."""
    if os.path.isfile(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, "r", encoding="utf-8") as cf:
                return json.load(cf)
        except:
            pass
    return {}

def save_config(data):
    """Save configuration to a local JSON file."""
    with open(CONFIG_FILE, "w", encoding="utf-8") as cf:
        json.dump(data, cf, indent=2)

def convert_repo_to_word(repo_path, output_docx="output.docx", excluded_extensions=None, excluded_directories=None):
    """Create a markdown file from a repo, then convert it to a Word doc with syntax highlighting."""
    temp_md = "temp_output.md"
    markdown_content = create_markdown_from_repo(
        repo_path,
        excluded_extensions=excluded_extensions or [],
        excluded_directories=excluded_directories or []
    )

    with open(temp_md, "w", encoding="utf-8") as f:
        f.write(markdown_content)

    convert_markdown_to_docx(temp_md, output_docx)
    print(f"Markdown to Word conversion complete: {output_docx}")
    os.remove(temp_md)  # Clean up temporary file

def main():
    config = load_config()
    print("Load existing config? (y/n) [y] ", end="")
    use_config = input().strip().lower() or "y"
    if use_config != "y":
        config = {}

    default_repo_path = config.get("repo_path", "")
    default_excl_ext = ",".join(config.get("excluded_extensions", []))
    default_excl_dirs = ",".join(config.get("excluded_directories", [".git"]))
    default_outfile = config.get("markdown_out", "output.md")
    default_docx_out = config.get("docx_out", "output.docx")

    repo_path = input(f"Enter the path to the repository [{default_repo_path}]: ").strip() or default_repo_path
    if not os.path.isdir(repo_path):
        print("Invalid path.")
        return

    excluded_str = input(f"Enter file extensions to exclude (comma-separated) [{default_excl_ext}]: ").strip()
    excluded_extensions = (
        [ext.strip() for ext in excluded_str.split(",") if ext.strip()] 
        if excluded_str else config.get("excluded_extensions", [])
    )

    dir_str = input(f"Enter directories to exclude (comma-separated) [{default_excl_dirs}]: ").strip()
    excluded_directories = (
        [d.strip() for d in dir_str.split(",") if d.strip()] 
        if dir_str else config.get("excluded_directories", [".git"])
    )

    print("Generate a Markdown file, then convert it to Word.")
    confirm = input(f"Proceed with '{repo_path}'? (y/n) [y]: ").strip().lower() or "y"
    if confirm != "y":
        print("Aborted.")
        return

    markdown_text = create_markdown_from_repo(
        repo_path,
        excluded_extensions=excluded_extensions,
        excluded_directories=excluded_directories
    )

    md_out = input(f"Markdown output file [{default_outfile}]: ").strip() or default_outfile
    with open(md_out, "w", encoding="utf-8") as out_f:
        out_f.write(markdown_text)

    print(f"Markdown generated: {md_out}")

    docx_out = input(f"Word output file [{default_docx_out}]: ").strip() or default_docx_out
    convert_markdown_to_docx(md_out, docx_out)

    print(f"Word document generated: {docx_out}")

    new_config = {
        "repo_path": repo_path,
        "excluded_extensions": excluded_extensions,
        "excluded_directories": excluded_directories,
        "markdown_out": md_out,
        "docx_out": docx_out
    }
    save_config(new_config)

if __name__ == "__main__":
    main()