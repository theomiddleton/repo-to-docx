import os
import json

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
    strip_trailing_newlines(CONFIG_FILE)

def strip_trailing_newlines(file_path):
    """Remove trailing lines with only whitespace (spaces, tabs, or newlines) at the end of a file."""
    with open(file_path, "r+", encoding="utf-8") as file:
        content = file.read()
        stripped_content = content.rstrip()  # Remove all trailing whitespace-only lines
        file.seek(0)
        file.write(stripped_content)
        file.truncate()

def _get_syntax_for_extension(extension):
    """Return a language name for syntax highlighting based on file extension."""
    extension_map = {
        ".py": "python",
        ".js": "javascript",
        ".cjs": "javascript",
        ".ts": "typescript",
        ".tsx": "tsx",
        ".jsx": "jsx",
        ".java": "java",
        ".c": "c",
        ".cpp": "cpp",
        ".sh": "bash",
        ".html": "html",
        ".css": "css",
        ".json": "json",
        ".yaml": "yaml",
        ".yml": "yaml",
        ".md": "markdown"
    }
    return extension_map.get(extension.lower(), "")

def create_markdown_from_repo(repo_path, excluded_extensions=None, excluded_directories=None):
    """
    Walk through a repository and generate a single Markdown document with all files.
    Files with extensions in 'excluded_extensions' will be skipped.
    Directories in 'excluded_directories' will be skipped entirely.
    Ignores any '.env' file by default.
    Returns a string containing the generated Markdown.
    """
    if excluded_extensions is None:
        excluded_extensions = []
    if excluded_directories is None:
        excluded_directories = []

    markdown_lines = []
    for root, dirs, files in os.walk(repo_path):
        # Remove directories we want to skip
        for d in list(dirs):
            if d in excluded_directories:
                dirs.remove(d)

        for file_name in files:
            if file_name == ".env":  # Always ignore .env
                continue

            extension = os.path.splitext(file_name)[1]
            if extension in excluded_extensions:
                continue

            file_path = os.path.join(root, file_name)
            print(f"Processing {file_path}...")

            with open(file_path, "r", encoding="utf-8", errors="ignore") as f:
                content = f.read()

            relative_path = os.path.relpath(file_path, repo_path)
            lang = _get_syntax_for_extension(extension)
            markdown_lines.append(f"## {relative_path}\n")
            markdown_lines.append(f"```{lang}")
            markdown_lines.append(content)
            markdown_lines.append("```\n")

    return "\n".join(markdown_lines)

def main():
    config = load_config()
    print("Load existing config? (y/n) [y] ", end="")
    use_config = input().strip().lower() or "y"
    if use_config != "y":
        config = {}  # Discard existing config

    # Autofill from config if present
    default_target = config.get("target", "")
    default_excl_ext = ",".join(config.get("excluded_extensions", []))
    default_excl_dirs = ",".join(config.get("excluded_directories", [".git"]))
    default_outfile = config.get("output_file", "output.md")

    # Prompt user, falling back to config if blank
    target = input(f"Enter the path to the repository [{default_target}]: ").strip() or default_target
    if not os.path.isdir(target):
        print("Invalid path.")
        return

    excluded_str = input(f"Enter file extensions to exclude (comma-separated, e.g. .exe,.dll) [{default_excl_ext}]: ").strip()
    excluded_extensions = [ext.strip() for ext in excluded_str.split(",") if ext.strip()] if excluded_str else (config.get("excluded_extensions") or [])

    dir_str = input(f"Enter directories to exclude (comma-separated) [{default_excl_dirs}]: ").strip()
    excluded_directories = [d.strip() for d in dir_str.split(",") if d.strip()] if dir_str else (config.get("excluded_directories") or [".git"])

    confirm = input(f"Generate Markdown from '{target}' excluding {excluded_extensions}, ignoring {excluded_directories}? (y/n) [y]: ").strip().lower() or "y"
    if confirm != "y":
        print("Aborted.")
        return

    markdown_doc = create_markdown_from_repo(
        target,
        excluded_extensions=excluded_extensions,
        excluded_directories=excluded_directories
    )

    output_file = input(f"Enter the name of the output Markdown file (e.g. output.md) [{default_outfile}]: ").strip() or default_outfile
    with open(output_file, "w", encoding="utf-8") as out_f:
        out_f.write(markdown_doc)
    print(f"Markdown document generated: {output_file}")

    # Save config for next time
    new_config = {
        "target": target,
        "excluded_extensions": excluded_extensions,
        "excluded_directories": excluded_directories,
        "output_file": output_file
    }
    save_config(new_config)

if __name__ == "__main__":
    main()
