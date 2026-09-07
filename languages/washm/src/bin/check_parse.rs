//! Parse (and lower) every file given on the command line without running it.
//!
//! Prints one line per failing file and a summary grouped by the first line
//! of the error, so grammar/walker gaps can be ranked by how many files they
//! block. Exit status is the number of failures (capped at 255).
//!
//!     check_parse tests/bash/**/*.sh
//!     check_parse --verbose file.sh      # also prints the lowered AST

use std::collections::BTreeMap;

fn main() {
    let mut verbose = false;
    let mut files = Vec::new();
    for arg in std::env::args().skip(1) {
        if arg == "--verbose" || arg == "-v" {
            verbose = true;
        } else {
            files.push(arg);
        }
    }
    let mut failures: BTreeMap<String, Vec<String>> = BTreeMap::new();
    let mut ok = 0usize;
    for path in &files {
        let src = match std::fs::read_to_string(path) {
            Ok(s) => s,
            Err(e) => {
                failures
                    .entry(format!("read error: {e}"))
                    .or_default()
                    .push(path.clone());
                continue;
            }
        };
        match vybe_language_washm::parse(&src) {
            Ok(module) => {
                ok += 1;
                if verbose {
                    println!("== {path}\n{module:#?}");
                }
            }
            Err(e) => {
                let key = e.lines().next().unwrap_or("").to_string();
                let key = if key.starts_with("Parse error") {
                    // Keep the "expected …" part only; positions differ per file.
                    let detail = e
                        .lines()
                        .skip_while(|l| !l.trim_start().starts_with('='))
                        .map(|l| l.trim())
                        .collect::<Vec<_>>()
                        .join(" ");
                    format!("parse: {detail}")
                } else {
                    key
                };
                println!("FAIL {path}: {}", e.replace('\n', " | "));
                failures.entry(key).or_default().push(path.clone());
            }
        }
    }
    let total: usize = failures.values().map(|v| v.len()).sum();
    println!("\n{ok} ok, {total} failed");
    let mut groups: Vec<(&String, &Vec<String>)> = failures.iter().collect();
    groups.sort_by_key(|(_, v)| std::cmp::Reverse(v.len()));
    for (k, v) in groups {
        println!("{:5}  {}", v.len(), k);
        for p in v.iter().take(3) {
            println!("         e.g. {p}");
        }
    }
    std::process::exit(total.min(255) as i32);
}
