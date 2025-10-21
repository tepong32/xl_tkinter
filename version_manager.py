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


def build_changelog_entry(new_version, message):
    """Build changelog text for display or write."""
    date_str = datetime.now().strftime("%Y-%m-%d")
    if "\n" in message:
        return f"## [{new_version}] - {date_str}\n### 🚀 Added\n{message.strip()}\n\n"
    else:
        return f"## [{new_version}] - {date_str}\n### 🚀 Added\n- {message.strip()}\n\n"


def update_files(new_version, message, dry_run=False):
    """Safely update version and changelog with UTF-8 encoding."""
    changelog_entry = build_changelog_entry(new_version, message)

    if dry_run:
        print(Style.BRIGHT + Fore.MAGENTA + "\n🚀 Dry Run Preview (no files written):\n")
        print(Fore.CYAN + f"📦 VERSION would become:\n{new_version}\n")
        print(Fore.CYAN + "📝 CHANGELOG entry would be:\n" + Fore.RESET + changelog_entry)
        print(Fore.CYAN + "🪣 Git commit (simulated):\n" + Fore.RESET + f"\"{message.strip()} (v{new_version})\"\n")
        print(Fore.GREEN + "✅ Nothing written. Use without --dry-run to apply changes.\n")
        return

    # --- Actual file write flow ---
    VERSION_FILE.write_text(new_version.strip(), encoding="utf-8")

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
    print(Fore.GREEN + f"✅ Updated VERSION and CHANGELOG.md → v{new_version}")


def git_commit_and_tag(new_version, message, dry_run=False):
    """Commit and tag new version in git."""
    if dry_run:
        print(Fore.MAGENTA + "💡 Skipping git commit and tag (dry-run mode)\n")
        return
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
        print("  python version_manager.py \"Commit message\" [major|minor|patch] [--dry-run]")
        print("  python version_manager.py \"\"\"Multiline changelog here\"\"\" [major|minor|patch] [--dry-run]")
        return

    args = sys.argv[1:]
    dry_run = "--dry-run" in args

    # Remove the flag so parsing is clean
    args = [a for a in args if a != "--dry-run"]

    if not args:
        print(Fore.RED + "❌ No commit message provided.")
        return

    raw_input = args[0]
    if raw_input.startswith(('"""', "'''")) and raw_input.endswith(('"""', "'''")):
        message = raw_input.strip().strip('"').strip("'")
    else:
        message = raw_input

    bump_type = args[1] if len(args) > 1 else "patch"

    current_version = get_current_version()
    new_version = bump_version(current_version, bump_type)
    update_files(new_version, message, dry_run=dry_run)
    git_commit_and_tag(new_version, message, dry_run=dry_run)

    print(Style.BRIGHT + Fore.GREEN + f"\n🎉 Done! {'(Preview only)' if dry_run else ''} Released version v{new_version}\n")


if __name__ == "__main__":
    main()
