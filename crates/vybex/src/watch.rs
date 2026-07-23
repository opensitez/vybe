//! `--watch`: Phase-1 hot reload (see `fulldebugplan.md` §7). Run the program,
//! then re-run it whenever the entry source changes. Each run is a fresh
//! subprocess — a clean-slate reload where program state is intentionally not
//! preserved (stateful in-process reload is Phase 2). This is the "save → see
//! the new output" dev loop, and it works for both short scripts (re-run on
//! change) and long-running servers (kill + restart on change). It composes
//! with `--debug`: each reload starts a fresh debug session.
//!
//! Driver-level only — nothing here touches the VM or its hot path.

use std::path::{Path, PathBuf};
use std::process::Command;
use std::time::{Duration, SystemTime};

const POLL: Duration = Duration::from_millis(200);

/// Enter the watch loop. `entry` is the source file to watch; `child_args` are
/// the CLI args to hand each child run (the original args with `--watch`/`-W`
/// removed so the child runs a single normal pass). Diverges — Ctrl-C stops it.
pub fn run_watch(entry: PathBuf, child_args: Vec<String>) -> ! {
    let exe = std::env::current_exe().unwrap_or_else(|_| PathBuf::from("vybex"));
    eprintln!("── watch mode ── {} (Ctrl-C to stop)", entry.display());
    loop {
        let baseline = mtime(&entry);
        eprintln!("\n▶ running…");
        let mut child = match Command::new(&exe).args(&child_args).spawn() {
            Ok(c) => c,
            Err(e) => {
                eprintln!("watch: failed to launch child: {e}");
                std::process::exit(1);
            }
        };

        // Run until the file changes or the child exits on its own.
        let changed_during_run = loop {
            match child.try_wait() {
                Ok(Some(_)) => break false, // child finished
                Ok(None) => {}
                Err(_) => break false,
            }
            if mtime(&entry) != baseline {
                break true;
            }
            std::thread::sleep(POLL);
        };

        if changed_during_run {
            // Long-running program (e.g. a server): stop it and restart now.
            let _ = child.kill();
            let _ = child.wait();
            eprintln!("↻ change detected — reloading");
            continue;
        }

        // Short program: it already exited. Wait for the next edit, then re-run.
        eprintln!("● exited — waiting for changes…");
        while mtime(&entry) == baseline {
            std::thread::sleep(POLL);
        }
        eprintln!("↻ change detected — reloading");
    }
}

fn mtime(path: &Path) -> Option<SystemTime> {
    std::fs::metadata(path).and_then(|m| m.modified()).ok()
}
