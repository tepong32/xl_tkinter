import subprocess
from datetime import datetime
from pathlib import Path
from colorama import Fore, Style, init

init(autoreset=True)

VERSION_FILE = Path("VERSION")
CHANGELOG_FILE = Path("CHANGELOG.md")


# -------------------------------
# Utilities
# -------------------------------

def get_current_version():
    """Read current version or create default."""
    if not VERSION_FILE.exists():
        VERSION_FILE.write_text("0.0.0", encoding="utf-8")
        print(Fore.YELLOW + "🪄 VERSION file not found — created default 0.0.0")

    try:
        version = VERSION_FILE.read_text(encoding="utf-8").strip()
        parts = version.split(".")
        if len(parts) != 3 or not all(p.isdigit() for p in parts):
            raise ValueError
        return version
    except Exception:
        print(Fore.RED + "⚠️ Invalid VERSION format — resetting to 0.0.0")
        VERSION_FILE.write_text("0.0.0", encoding="utf-8")
        return "0.0.0"


def bump_version(current_version, bump_type):
    """Increment version based on bump type."""
    major, minor, patch = map(int, current_version.split("."))

    if bump_type == "major":
        major += 1
        minor = 0
        patch = 0
    elif bump_type == "minor":
        minor += 1
        patch = 0
    else:
        patch += 1

    new_version = f"{major}.{minor}.{patch}"
    print(f"{Fore.CYAN}🔧 Bumping version {current_version} → {new_version}")
    return new_version


def update_files(new_version, message):
    """Safely update version and changelog with UTF-8 encoding."""
    VERSION_FILE.write_text(new_version.strip(), encoding="utf-8")
    date_str = datetime.now().strftime("%Y-%m-%d")

    # Preserve formatting for multiline markdown messages
    if "\n" in message:
        changelog_entry = f"## [{new_version}] - {date_str}\n### 🚀 Added\n{message.strip()}\n\n"
    else:
        changelog_entry = f"## [{new_version}] - {date_str}\n### 🚀 Added\n- {message.strip()}\n\n"

    # Read or create changelog
    if CHANGELOG_FILE.exists():
        try:
            content = CHANGELOG_FILE.read_text(encoding="utf-8").strip()
        except UnicodeDecodeError:
            print(Fore.YELLOW + "⚠️ Changelog encoding issue — resetting to fresh format.")
            content = "# Changelog\n\n"
        except Exception as e:
            print(Fore.YELLOW + f"⚠️ Could not read changelog: {e}")
            content = "# Changelog\n\n"
    else:
        content = "# Changelog\n\n"
        print(Fore.YELLOW + "🪄 CHANGELOG.md not found — created fresh one")

    if not content.startswith("# Changelog"):
        content = "# Changelog\n\n" + content

    try:
        insert_index = content.find('\n', content.find('# Changelog') + len('# Changelog'))
        while insert_index + 1 < len(content) and content[insert_index + 1] in ('\n', ' '):
            insert_index += 1
        new_content = content[:insert_index] + "\n" + changelog_entry + content[insert_index:].strip()
    except Exception:
        new_content = content + "\n" + changelog_entry
        print(Fore.YELLOW + "⚠️ Failed to insert at top, appending to end.")

    CHANGELOG_FILE.write_text(new_content.strip() + "\n", encoding="utf-8")


def git_commit_and_tag(new_version, message):
    """Commit and tag new version in git."""
    try:
        subprocess.run(["git", "add", "VERSION", "CHANGELOG.md"], check=True)
        subprocess.run(["git", "commit", "-m", f"{message} (v{new_version})"], check=True)
        subprocess.run(["git", "tag", f"v{new_version}"], check=True)
        print(Fore.GREEN + f"✅ Git commit + tag created for v{new_version}")
    except subprocess.CalledProcessError:
        print(Fore.RED + "⚠️ Git commit or tag failed. Check if repo is clean.")


# -------------------------------
# Main CLI logic
# -------------------------------

def main():
    import sys

    if len(sys.argv) < 2:
        print(Fore.YELLOW + "Usage:")
        print("  python version_manager.py \"Commit message\" [major|minor|patch]")
        print("  python version_manager.py \"\"\"Multiline changelog here\"\"\" [major|minor|patch]")
        return

    # Handle multiline message (triple quotes)
    raw_input = sys.argv[1]
    if raw_input.startswith(('"""', "'''")) and raw_input.endswith(('"""', "'''")):
        message = raw_input.strip().strip('"').strip("'")
    else:
        message = raw_input

    bump_type = sys.argv[2] if len(sys.argv) > 2 else "patch"

    current_version = get_current_version()
    new_version = bump_version(current_version, bump_type)
    update_files(new_version, message)
    git_commit_and_tag(new_version, message)

    print(Style.BRIGHT + Fore.GREEN + f"\n🎉 Done! Released version v{new_version}\n")


if __name__ == "__main__":
    main()
