//! Execute one extracted test file and turn it into a verdict.
//!
//! The verdict is the exit code and nothing else — no stdout parsing. `vybex`
//! already exits 1 on an uncaught error and 0 on a clean run, so this needs no
//! runtime change, and it is immune to the `[vybex] Project …` banner that
//! compilation prints on stdout before the program starts.
//!
//! Each test carries its harness inline (spliced from `harness/<lang>/`), so a
//! case is one self-contained file — which is what lets the same path go to
//! `vybex`, to `go run`, or into the step debugger unchanged.

use std::path::Path;
use std::process::Command;

#[derive(Clone, Copy, PartialEq, Eq, Debug)]
pub enum Mode {
    Run,
    Compile,
    CompileFail,
}

impl Mode {
    /// Read the `// vybe-test-mode:` directive out of a file's header.
    pub fn of(text: &str) -> Mode {
        for line in text.lines().take(10) {
            if let Some((_, rest)) = line.split_once("vybe-test-mode:") {
                return match rest.trim() {
                    "compile" => Mode::Compile,
                    "compile-fail" => Mode::CompileFail,
                    _ => Mode::Run,
                };
            }
        }
        Mode::Run
    }
}

pub struct Outcome {
    pub pass: bool,
    pub code: Option<i32>,
    pub output: String,
}

pub fn run_case(vybex: &Path, file: &Path, mode: Mode) -> Outcome {
    let mut cmd = Command::new(vybex);
    if mode != Mode::Run {
        // `-d` disassembles without running: the frontend must accept the
        // program, nothing more. That is exactly what `compile_ok` asserted.
        cmd.arg("-d");
    }
    cmd.arg(file);
    execute(cmd, mode)
}

/// Run the same files under a foreign runtime — `go run`, `python3`, `node`,
/// `php`. This is the whole reason the tests are ordinary source: the
/// expectations were written against Vybe's behaviour, so a diff here is
/// evidence about the corpus, not only about us.
pub fn run_foreign(program: &str, args: &[String], file: &Path, mode: Mode) -> Outcome {
    let mut cmd = Command::new(program);
    cmd.args(args).arg(file);
    execute(cmd, mode)
}

fn execute(mut cmd: Command, mode: Mode) -> Outcome {
    match cmd.output() {
        Ok(out) => {
            let clean = out.status.success();
            let mut text = String::from_utf8_lossy(&out.stdout).into_owned();
            text.push_str(&String::from_utf8_lossy(&out.stderr));
            Outcome {
                pass: if mode == Mode::CompileFail { !clean } else { clean },
                code: out.status.code(),
                output: text,
            }
        }
        Err(e) => Outcome { pass: false, code: None, output: format!("spawn failed: {e}") },
    }
}

/// The first line that explains a failure. The harness prints its own
/// `FAIL: want [...] got [...]`, which is the useful one when present;
/// otherwise fall back to the runtime's own error line.
pub fn failure_line(output: &str) -> String {
    output
        .lines()
        .find(|l| l.trim_start().starts_with("FAIL: "))
        .or_else(|| {
            output
                .lines()
                .find(|l| l.contains("rror") || l.contains("panic:"))
        })
        .unwrap_or("(no output)")
        .trim()
        .to_string()
}
