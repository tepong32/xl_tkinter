#!/usr/bin/env python3
"""
version_manager.py
Automates version bumping, changelog updates, git tagging, and commits.
Usage:
    python version_manager.py "Your changelog message" [major|minor|patch]
"""

import subprocess
import sys
from datetime import datetime
from pathlib import Path

# --- CONFIG ---
VERSION_FILE = Path("VERSION")
CHANGELOG_FILE = Path("CHANGELOG.md")
DEFAULT_BUMP = "patch"  # default if not specified
# ----------------

def run(cmd):
    """Run shell commands safely."""
    result = subprocess.run(cmd, shell=True, capture_output=True, text=True)
    if result.returncode != 0:
        print("❌ Error:", result.stderr.strip())
        sys.exit(1)
    return result.stdout.strip()

def get_current_version():
    if not VERSION_FILE.exists():
        VERSION_FILE.write_text("0.0.0")
    return VERSION_FILE.read_text().strip()

def bump_version(current_version, bump_type):
    major, minor, patch = map(int, current_version.split("."))
    if bump_type == "major":
        major += 1; minor = 0; patch = 0
    elif bump_type == "minor":
        minor += 1; patch = 0
    else:  # patch
        patch += 1
    return f"{major}.{minor}.{patch}"

def update_files(new_version, message):
    """Update version and changelog."""
    VERSION_FILE.write_text(new_version, encoding="utf-8")
    date_str = datetime.now().strftime("%Y-%m-%d")

    changelog_entry = f"## v{new_version} - {date_str}\n- {message}\n\n"
    if CHANGELOG_FILE.exists():
        try:
            content = CHANGELOG_FILE.read_text(encoding="utf-8")
        except UnicodeDecodeError:
            content = CHANGELOG_FILE.read_text(encoding="utf-8", errors="ignore")
    else:
        content = "# Changelog\n\n"

    CHANGELOG_FILE.write_text(content + changelog_entry, encoding="utf-8")
    
def git_commit_and_tag(new_version, message):
    """Commit changes and push with tag."""
    run("git add VERSION CHANGELOG.md")
    run(f'git commit -m "chore(release): v{new_version} - {message}"')
    run(f'git tag -a v{new_version} -m "Release v{new_version}"')
    run("git push")
    run("git push --tags")

def main():
    if len(sys.argv) < 2:
        print("Usage: python version_manager.py \"Message\" [major|minor|patch]")
        sys.exit(1)

    message = sys.argv[1]
    bump_type = sys.argv[2] if len(sys.argv) > 2 else DEFAULT_BUMP

    current_version = get_current_version()
    new_version = bump_version(current_version, bump_type)

    print(f"🔧 Bumping version {current_version} → {new_version}")
    update_files(new_version, message)
    git_commit_and_tag(new_version, message)

    print(f"✅ Done! Released version v{new_version}")

if __name__ == "__main__":
    main()
