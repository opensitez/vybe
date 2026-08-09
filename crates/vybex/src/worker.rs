//! Warm execution mode — boot once, run many programs without relaunching.
//!
//! `vybex --worker` boots a VM, registers every plugin and platform, snapshots
//! the warm baseline, and then reads one job per line from stdin. Each job runs
//! against a `reset_to`'d VM: the previous program's globals, heap, linear
//! memory, tables and appended chunks are rolled back to boot.
//!
//! This is not a test feature. Roughly 90% of a `vybex <file>` invocation is
//! setup — an empty program costs 0.204s against 0.019s of process spawn — so
//! anything that runs many short programs pays that over and over: a test
//! suite, a request handler, a serverless invocation. `VM::reset_to` was
//! already written with the last of those in mind; its own comment says "no
//! earlier tenant's bytes survive".
//!
//! ## Protocol
//!
//! Line-oriented, on stdout:
//!
//! ```text
//! ← ##vybe-ready                       once, after boot
//! → /path/to/program.go                one job per line, optional \tcompile
//! ← ...the program's own output...
//! ← ##vybe-result\tok                  or  ##vybe-result\terr\t<message>
//! → /path/to/harness.js\tpreload       fold into the baseline (see below)
//! ```
//!
//! ## Loaded state vs tenant state
//!
//! Two kinds of code run here and they are not the same thing:
//!
//! - a **tenant** — one job, one program. It must leave NOTHING behind: the
//!   next job has to be unable to tell it ever ran. That is `reset_to`'s
//!   contract, and any difference between warm and cold execution is a bug in
//!   it, not a property of warm mode.
//! - **loaded state** — a test harness, a runtime prelude: code the EMBEDDER
//!   puts there deliberately, that every job is supposed to see. Sent with
//!   mode `preload`, it runs into the VM and the baseline is re-snapshotted,
//!   so it costs one load per worker rather than one per job.
//!
//! Without the second, the only way to make code persist was for it to survive
//! a reset — which is precisely the property a tenant must never have. Keeping
//! them separate is what lets the reset be total.
//!
//! The program's output is NOT captured or redirected — it goes to this
//! process's real stdout through the real host functions, and the sentinel line
//! terminates it. That is deliberate: capturing would mean registering
//! substitute output handlers, and a substitute that drifts from the real one
//! is exactly how `wasi:logging/logging.log` came to behave differently under
//! the test harness than under vybex.

use std::io::{BufRead, Write};
use vybe_runtime::VM;
use vybe_runtime::capabilities::Capabilities;

pub const READY: &str = "##vybe-ready";
pub const RESULT: &str = "##vybe-result";

pub fn run(caps: Capabilities) -> ! {
    // The boot sequence — tracking on before the first allocation, plugins,
    // adapters, primed prototypes, snapshot — lives in `crate::warm` because
    // `--serve` needs the identical one. When it lived here in a copy, the
    // server's copy was missing two of those steps.
    let (mut vm, mut baseline) = match crate::warm::boot(&caps) {
        Ok(pair) => pair,
        Err(e) => {
            eprintln!("[worker] {e}");
            std::process::exit(1);
        }
    };

    println!("{READY}");
    let _ = std::io::stdout().flush();

    let stdin = std::io::stdin();
    for line in stdin.lock().lines().map_while(Result::ok) {
        let line = line.trim_end();
        if line.is_empty() {
            continue;
        }
        // `<path>\t<mode>\t<expected-exit>` — the third field is optional so an
        // older driver still works.
        let mut fields = line.split('\t');
        let path = fields.next().unwrap_or(line);
        let mode = fields.next().unwrap_or("run");
        let want_exit: i32 = fields
            .next()
            .and_then(|f| f.trim().parse().ok())
            .unwrap_or(0);

        // `preload` is LOADED STATE, not a tenant. The file is run into the warm
        // VM and the baseline is re-snapshotted, so every later reset restores
        // TO it: a test harness or a runtime prelude is paid for ONCE per
        // worker instead of once per job, without ever becoming something a
        // reset has to flush.
        //
        // The distinction is the whole point. A job's script must leave nothing
        // behind — that is what `reset_to` guarantees. Something the EMBEDDER
        // loads on purpose is not the script, and there was previously no way
        // to say so: the only way to make code survive was to have it survive a
        // reset, which is exactly the property a tenant must not have.
        if mode == "preload" {
            crate::warm::reset(&mut vm, &baseline);
            match run_job(&mut vm, path, "run", &caps) {
                Ok(()) => {
                    baseline = vm.snapshot();
                    println!("{RESULT}\tok");
                }
                // A preload that fails leaves the OLD baseline in place — a
                // half-loaded harness would poison every job after it, which is
                // the failure this whole mechanism exists to prevent.
                Err(e) => {
                    crate::warm::reset(&mut vm, &baseline);
                    println!("{RESULT}\terr\t{}", one_line(&e));
                }
            }
            let _ = std::io::stdout().flush();
            continue;
        }

        crate::warm::reset(&mut vm, &baseline);

        let outcome = run_job(&mut vm, path, mode, &caps);
        let _ = std::io::stdout().flush();
        // The verdict is the EXIT CODE, so a program that called
        // `wasi:cli/exit` with a non-zero status failed even though `run`
        // returned `Ok`. Reading only the `Result` made warm mode disagree
        // with `vybex <file>`, with `--cold`, and with the real runtime: a
        // Python `sys.exit(3)` passed under the default pool and failed
        // everywhere else. `--cold` exists to catch exactly this.
        match outcome {
            Ok(()) if vm.pending_exit_code != want_exit => println!(
                "{RESULT}\terr\texited with status {} (expected {want_exit})",
                vm.pending_exit_code
            ),
            Ok(()) => println!("{RESULT}\tok"),
            Err(e) => println!("{RESULT}\terr\t{}", one_line(&e)),
        }
        let _ = std::io::stdout().flush();
    }
    std::process::exit(0)
}

fn run_job(vm: &mut VM, path: &str, mode: &str, caps: &Capabilities) -> Result<(), String> {
    let paths = [std::path::PathBuf::from(path)];
    let program = vybe_compiler::projects::load_program(&paths)?;
    let mut units = program.units;
    let entry = units.pop().ok_or("empty program")?;

    let mut compiler = crate::dynamic::RuntimeCompilerService::with_capabilities(vm, caps.clone());

    // `compile` mode asks only that the front-end accepts the program — the
    // same assertion the Rust `compile_ok` helpers made.
    if mode == "compile" || mode == "compile-fail" {
        let result = compiler.compile_bundle(&entry).map(|_| ());
        return match (mode, &result) {
            ("compile-fail", Ok(())) => Err("expected the front-end to reject this".into()),
            ("compile-fail", Err(_)) => Ok(()),
            _ => result,
        };
    }

    // Other languages of a multi-language program link in first, exactly as on
    // the ordinary run path.
    for unit in &units {
        compiler.run_program_unit(unit)?;
    }
    compiler.compile_and_run_bundle(&entry).map(|_| ())
}

/// Collapse a multi-line error into one protocol line.
fn one_line(text: &str) -> String {
    text.lines()
        .map(str::trim)
        .filter(|l| !l.is_empty())
        .collect::<Vec<_>>()
        .join(" · ")
        .chars()
        .take(400)
        .collect()
}
