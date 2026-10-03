//! Standalone grammar compilation driver; no language/compiler/VM linkage.
use std::{env, fs, process::ExitCode, time::Instant};

fn main() -> ExitCode {
    let paths: Vec<_> = env::args_os().skip(1).collect();
    if paths.is_empty() {
        eprintln!(
            "usage: grammarcheck <grammar.pest> [grammar.pest ...]\nChecks grammar syntax, reference resolution, recursion and progress."
        );
        return ExitCode::from(2);
    }
    let mut failed = false;
    for path in paths {
        let label = path.to_string_lossy();
        let source = match fs::read_to_string(&path) {
            Ok(source) => source,
            Err(error) => {
                eprintln!("{label}: {error}");
                failed = true;
                continue;
            }
        };
        let started = Instant::now();
        match vybe_parser::compile(&source) {
            Ok(grammar) => {
                for warning in &grammar.analysis().warnings {
                    eprintln!("{}", warning.render(&label, &source));
                }
                println!(
                    "{label}: {} rules, {} expressions, {:.3} ms (syntax + resolution + analysis)",
                    grammar.syntax().rules.len(),
                    grammar.syntax().expressions.len(),
                    started.elapsed().as_secs_f64() * 1000.0
                );
            }
            Err(errors) => {
                failed = true;
                for error in errors {
                    eprintln!("{}", error.render(&label, &source));
                    for (span, message) in error.related {
                        let related = vybe_parser::Diagnostic {
                            code: "note",
                            message,
                            span,
                            related: Vec::new(),
                        };
                        eprintln!("{}", related.render(&label, &source));
                    }
                }
            }
        }
    }
    if failed {
        ExitCode::FAILURE
    } else {
        ExitCode::SUCCESS
    }
}
