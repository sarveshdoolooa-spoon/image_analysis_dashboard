#!/usr/bin/env python3
"""
Batch commit and push files safely.

Example:
  python batch_git_upload.py --dry-run
  python batch_git_upload.py --batch-size-mb 50 --push
  python batch_git_upload.py --batch-size-mb 100 --push --branch main
"""

from __future__ import annotations

import argparse
import subprocess
import sys
from pathlib import Path
from typing import Iterable


DEFAULT_EXCLUDE_DIRS = {
    ".git",
    "node_modules",
    "venv",
    ".venv",
    "__pycache__",
    ".pytest_cache",
    ".mypy_cache",
    ".ruff_cache",
    "dist",
    "build",
}

DEFAULT_EXCLUDE_FILES = {
    ".DS_Store",
    "Thumbs.db",
}


def run(cmd: list[str], check: bool = True) -> subprocess.CompletedProcess:
    print(f"> {' '.join(cmd)}")
    return subprocess.run(cmd, text=True, check=check)


def get_repo_root() -> Path:
    result = subprocess.run(
        ["git", "rev-parse", "--show-toplevel"],
        text=True,
        capture_output=True,
    )
    if result.returncode != 0:
        print("Error: this script must be run inside a Git repository.")
        sys.exit(1)

    return Path(result.stdout.strip()).resolve()


def ensure_clean_index() -> None:
    result = subprocess.run(
        ["git", "diff", "--cached", "--quiet"],
        text=True,
    )
    if result.returncode != 0:
        print("Error: you already have staged files.")
        print("Run `git status`, then commit/reset them before using this script.")
        sys.exit(1)


def file_size_mb(path: Path) -> float:
    return path.stat().st_size / (1024 * 1024)


def should_skip(
    path: Path,
    exclude_dirs: set[str],
    exclude_files: set[str],
    max_file_mb: float,
) -> bool:
    if any(part in exclude_dirs for part in path.parts):
        return True

    if path.name in exclude_files:
        return True

    if file_size_mb(path) > max_file_mb:
        return True

    return False


def collect_files(
    repo_root: Path,
    exclude_dirs: set[str],
    exclude_files: set[str],
    max_file_mb: float,
) -> list[Path]:
    print("Loading tracked files...")
    tracked_result = subprocess.run(
        ["git", "ls-files"],
        cwd=repo_root,
        text=True,
        capture_output=True,
        check=True,
    )

    tracked_files = {
        Path(line.strip())
        for line in tracked_result.stdout.splitlines()
        if line.strip()
    }

    print("Scanning repository files...")
    files: list[Path] = []
    scanned = 0

    for path in repo_root.rglob("*"):
        if not path.is_file():
            continue

        scanned += 1

        if scanned % 1000 == 0:
            print(f"Scanned {scanned} files, eligible {len(files)}...")

        rel = path.relative_to(repo_root)

        if rel in tracked_files:
            continue

        if should_skip(rel, exclude_dirs, exclude_files, max_file_mb):
            continue

        files.append(rel)

    files.sort(key=lambda p: str(p))
    print(f"Scan complete. Scanned {scanned} files.")
    return files


def make_batches(files: Iterable[Path], repo_root: Path, batch_size_mb: float) -> list[list[Path]]:
    batches: list[list[Path]] = []
    current: list[Path] = []
    current_size = 0.0

    for rel_path in files:
        size = file_size_mb(repo_root / rel_path)

        if current and current_size + size > batch_size_mb:
            batches.append(current)
            current = []
            current_size = 0.0

        current.append(rel_path)
        current_size += size

    if current:
        batches.append(current)

    return batches


def commit_batch(
    batch: list[Path],
    repo_root: Path,
    batch_number: int,
    total_batches: int,
    message_prefix: str,
    dry_run: bool,
) -> bool:
    total_mb = sum(file_size_mb(repo_root / p) for p in batch)

    print()
    print(f"Batch {batch_number}/{total_batches}")
    print(f"Files: {len(batch)}")
    print(f"Size: {total_mb:.2f} MB")

    if dry_run:
        for path in batch[:20]:
            print(f"  {path}")
        if len(batch) > 20:
            print(f"  ... and {len(batch) - 20} more")
        return False

    for path in batch:
        run(["git", "add", "--", str(path)])

    result = subprocess.run(["git", "diff", "--cached", "--quiet"])

    if result.returncode == 0:
        print("No staged changes in this batch. Skipping commit.")
        return False

    run([
        "git",
        "commit",
        "-m",
        f"{message_prefix} {batch_number}/{total_batches}",
    ])

    return True


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Commit and optionally push repository files in safe batches."
    )

    parser.add_argument("--batch-size-mb", type=float, default=50)
    parser.add_argument("--max-file-mb", type=float, default=1)
    parser.add_argument("--remote", default="origin")
    parser.add_argument("--branch", default="main")
    parser.add_argument("--push", action="store_true")
    parser.add_argument("--dry-run", action="store_true")
    parser.add_argument("--message-prefix", default="Batch upload")
    parser.add_argument("--exclude-dir", action="append", default=[])
    parser.add_argument("--exclude-file", action="append", default=[])

    args = parser.parse_args()

    repo_root = get_repo_root()
    print(f"Repo: {repo_root}")

    ensure_clean_index()

    exclude_dirs = DEFAULT_EXCLUDE_DIRS | set(args.exclude_dir)
    exclude_files = DEFAULT_EXCLUDE_FILES | set(args.exclude_file)

    files = collect_files(
        repo_root=repo_root,
        exclude_dirs=exclude_dirs,
        exclude_files=exclude_files,
        max_file_mb=args.max_file_mb,
    )

    if not files:
        print("No files found to commit.")
        return

    batches = make_batches(files, repo_root, args.batch_size_mb)

    print(f"Eligible files: {len(files)}")
    print(f"Total batches: {len(batches)}")
    print(f"Batch size: {args.batch_size_mb} MB")
    print(f"Max file size: {args.max_file_mb} MB")

    committed_any = False

    for index, batch in enumerate(batches, start=1):
        committed = commit_batch(
            batch=batch,
            repo_root=repo_root,
            batch_number=index,
            total_batches=len(batches),
            message_prefix=args.message_prefix,
            dry_run=args.dry_run,
        )

        committed_any = committed_any or committed

        if committed and args.push:
            run(["git", "push", args.remote, args.branch])

    if args.dry_run:
        print()
        print("Dry run complete. No files were committed.")
        print("Run without --dry-run to commit.")
        return

    if committed_any and not args.push:
        print()
        print("Commits created locally.")
        print(f"Push with: git push {args.remote} {args.branch}")

    print("Done.")


if __name__ == "__main__":
    main()