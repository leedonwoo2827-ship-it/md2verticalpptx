"""
md2pptx CLI — Build PowerPoint presentations from extended markdown.

Usage:
    python -m md2pptx <body.md> -t <templates_dir> [-o output.pptx] [options]
"""

from __future__ import annotations

import argparse
import io
import os
import sys
import time
from pathlib import Path

# Fix Windows cp949 encoding issues with Unicode characters
if sys.platform == 'win32':
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8', errors='replace')

from rich.console import Console
from rich.panel import Panel
from rich.progress import Progress, SpinnerColumn, BarColumn, TextColumn, MofNCompleteColumn, TimeElapsedColumn
from rich.table import Table

from . import __version__
from .models import SlideData, SlideResult, BuildSummary
from .parser import parse_md
from .builder import load_slide_index, build_single_slide, merge_pptx_files

console = Console()


def parse_args(argv: list[str] | None = None) -> argparse.Namespace:
    p = argparse.ArgumentParser(
        prog='md2pptx',
        description='Build PowerPoint presentations from extended markdown.',
    )
    p.add_argument('input_md', help='Extended markdown body file (proposal-body-partN.md)')
    p.add_argument('-t', '--templates', required=True, help='Template directory with S*.pptx and slide_index.json')
    p.add_argument('-o', '--output', default='', help='Output PPTX path (default: output/result.pptx)')
    p.add_argument('--slides-dir', default='', help='Directory for individual slide PPTX files')
    p.add_argument('--batch-size', type=int, default=25, help='Merge batch size (default: 25)')
    p.add_argument('--keep-slides', action='store_true', help='Keep individual slide files after merge')
    p.add_argument('--no-merge', action='store_true', help='Build individual slides only, skip merge')
    p.add_argument('--no-notes', action='store_true', help='Skip writing slide notes')
    p.add_argument('--continue-on-error', action='store_true', help='Skip failed slides instead of aborting')
    p.add_argument('-v', '--verbose', action='store_true', help='Show detailed output')
    p.add_argument('-q', '--quiet', action='store_true', help='Minimal output')
    p.add_argument('--version', action='version', version=f'%(prog)s {__version__}')
    return p.parse_args(argv)


def resolve_paths(args: argparse.Namespace) -> tuple[Path, Path, Path, Path]:
    """Resolve and validate all paths. Returns (input_md, templates_dir, output_path, slides_dir)."""
    input_md = Path(args.input_md)
    if not input_md.exists():
        console.print(f'[red]Error:[/] Input file not found: {input_md}')
        sys.exit(1)

    templates_dir = Path(args.templates)
    if not templates_dir.is_dir():
        console.print(f'[red]Error:[/] Templates directory not found: {templates_dir}')
        sys.exit(1)

    idx_path = templates_dir / 'slide_index.json'
    if not idx_path.exists():
        console.print(f'[red]Error:[/] slide_index.json not found in {templates_dir}')
        sys.exit(1)

    # Output path
    if args.output:
        output_path = Path(args.output)
    else:
        output_path = input_md.parent / 'output' / 'result.pptx'

    output_path.parent.mkdir(parents=True, exist_ok=True)

    # Slides directory
    if args.slides_dir:
        slides_dir = Path(args.slides_dir)
    else:
        slides_dir = output_path.parent / 'slides'

    slides_dir.mkdir(parents=True, exist_ok=True)

    return input_md, templates_dir, output_path, slides_dir


def build_all_slides(
    slides: list[SlideData],
    slides_info: dict,
    templates_dir: Path,
    slides_dir: Path,
    write_notes: bool = True,
    continue_on_error: bool = False,
    verbose: bool = False,
    quiet: bool = False,
) -> list[SlideResult]:
    """Build all individual slide PPTX files with progress display."""
    results: list[SlideResult] = []
    total = len(slides)
    succeeded = 0
    failed = 0

    with Progress(
        SpinnerColumn(),
        TextColumn('[bold blue]{task.description}'),
        BarColumn(),
        MofNCompleteColumn(),
        TextColumn('·'),
        TimeElapsedColumn(),
        console=console,
        disable=quiet,
    ) as progress:
        task = progress.add_task('Building slides', total=total)

        for slide in slides:
            out_path = slides_dir / f'slide_{slide.index:03d}.pptx'
            t0 = time.time()

            ref_label = f'S{slide.ref_slide:04d}' if slide.ref_slide else '????'
            tmpl_label = slide.template or '??'
            progress.update(task, description=f'Building slides  [dim][{ref_label}] {tmpl_label}[/]')

            try:
                build_single_slide(
                    slide=slide,
                    slides_info=slides_info,
                    output_path=str(out_path),
                    slides_dir=str(templates_dir),
                    write_notes=write_notes,
                )
                elapsed = time.time() - t0
                results.append(SlideResult(
                    index=slide.index,
                    ref_slide=slide.ref_slide,
                    template=slide.template,
                    status='success',
                    output_path=str(out_path),
                    elapsed=elapsed,
                ))
                succeeded += 1

                if verbose and not quiet:
                    console.print(f'  [green]✓[/] #{slide.index:03d} [{ref_label}] {tmpl_label} ({elapsed:.1f}s)')

            except Exception as e:
                elapsed = time.time() - t0
                results.append(SlideResult(
                    index=slide.index,
                    ref_slide=slide.ref_slide,
                    template=slide.template,
                    status='failed',
                    error=str(e),
                    elapsed=elapsed,
                ))
                failed += 1

                if not quiet:
                    console.print(f'  [red]✗[/] #{slide.index:03d} [{ref_label}] {e}')

                if not continue_on_error:
                    console.print('[red]Aborting.[/] Use --continue-on-error to skip failures.')
                    break

            progress.update(task, advance=1)

    if not quiet:
        console.print(f'  [green]✓ {succeeded}[/] succeeded  [red]✗ {failed}[/] failed  / {total} total')

    return results


def merge_slides_cli(
    results: list[SlideResult],
    output_path: Path,
    batch_size: int = 25,
    quiet: bool = False,
) -> None:
    """Merge successful slides into final PPTX with progress."""
    successful = [r for r in results if r.status == 'success']
    if not successful:
        console.print('[red]No successful slides to merge.[/]')
        return

    files = [r.output_path for r in successful]

    with Progress(
        SpinnerColumn(),
        TextColumn('[bold blue]{task.description}'),
        BarColumn(),
        MofNCompleteColumn(),
        TextColumn('·'),
        TimeElapsedColumn(),
        console=console,
        disable=quiet,
    ) as progress:
        task = progress.add_task('Merging slides', total=len(files) - 1 if len(files) > 1 else 1)

        def on_progress(current: int, total: int):
            progress.update(task, completed=current)

        merge_pptx_files(
            part_files=files,
            output_path=str(output_path),
            batch_size=batch_size,
            on_progress=on_progress,
        )

        progress.update(task, completed=progress.tasks[task].total)


def print_summary(summary: BuildSummary, failed_results: list[SlideResult]) -> None:
    """Print final summary panel."""
    if summary.output_path and os.path.exists(summary.output_path):
        size = os.path.getsize(summary.output_path)
        size_str = f'{size / 1024 / 1024:.1f} MB' if size > 1024 * 1024 else f'{size / 1024:.0f} KB'
    else:
        size_str = 'N/A'

    minutes = int(summary.total_time // 60)
    seconds = int(summary.total_time % 60)
    time_str = f'{minutes}m {seconds}s' if minutes else f'{seconds}s'

    status = '[green]✓ Build complete[/]' if summary.failed == 0 else '[yellow]⚠ Build complete (with errors)[/]'

    lines = [
        status,
        f'Slides: {summary.succeeded}/{summary.total}' + (f' ([red]{summary.failed} failed[/])' if summary.failed else ''),
    ]

    if summary.output_path:
        lines.append(f'Output: {summary.output_path}')
        lines.append(f'Size:   {size_str}')

    lines.append(f'Time:   {time_str}')

    if failed_results:
        lines.append('')
        lines.append('[red]Failed slides:[/]')
        for r in failed_results[:10]:
            ref = f'S{r.ref_slide:04d}' if r.ref_slide else '????'
            lines.append(f'  #{r.index:03d} [{ref}] {r.error}')
        if len(failed_results) > 10:
            lines.append(f'  ... and {len(failed_results) - 10} more')

    console.print(Panel('\n'.join(lines), title='md2pptx', border_style='blue'))


def main(argv: list[str] | None = None) -> None:
    args = parse_args(argv)
    input_md, templates_dir, output_path, slides_dir = resolve_paths(args)

    t_start = time.time()

    # 1. Parse
    if not args.quiet:
        console.print(f'Parsing [cyan]{input_md.name}[/]...', end=' ')

    md_text = input_md.read_text(encoding='utf-8')
    config, slides = parse_md(md_text)

    if not slides:
        console.print('[red]No slides found in input file.[/]')
        sys.exit(1)

    if not args.quiet:
        console.print(f'[green]{len(slides)}[/] slides found')

    # 2. Load slide index
    idx_path = templates_dir / 'slide_index.json'
    slides_info = load_slide_index(str(idx_path))

    # 3. Build individual slides
    results = build_all_slides(
        slides=slides,
        slides_info=slides_info,
        templates_dir=templates_dir,
        slides_dir=slides_dir,
        write_notes=not args.no_notes,
        continue_on_error=args.continue_on_error,
        verbose=args.verbose,
        quiet=args.quiet,
    )

    # 4. Merge
    if not args.no_merge:
        merge_slides_cli(results, output_path, batch_size=args.batch_size, quiet=args.quiet)

    # 5. Cleanup individual slides (unless --keep-slides or --no-merge)
    if not args.keep_slides and not args.no_merge:
        for r in results:
            if r.status == 'success' and r.output_path and os.path.exists(r.output_path):
                try:
                    os.remove(r.output_path)
                except Exception:
                    pass

    # 6. Summary
    total_time = time.time() - t_start
    failed_results = [r for r in results if r.status == 'failed']
    summary = BuildSummary(
        results=results,
        output_path=str(output_path) if not args.no_merge else '',
        total_time=total_time,
    )

    if not args.quiet:
        console.print()
        print_summary(summary, failed_results)

    sys.exit(1 if failed_results else 0)
