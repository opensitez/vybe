//! PHP emitter: one extracted case → a standalone `.php` test.
//!
//! Unlike Go and JS this does not rewrite output calls. `echo` is a statement
//! with no implicit newline, so consecutive echos share a line and there is no
//! i-th print to pair with an i-th expected line. The program instead runs
//! inside an output buffer and its whole output is checked once — exact, and
//! independent of how the program chose to write it.

use crate::emit::go::Pairing;
use crate::extract::Case;

pub struct Emitted {
    pub text: String,
    pub pairing: Pairing,
}

pub fn emit(case: &Case, origin: &str, slug: &str, harness: &str) -> Emitted {
    let header = format!("// vybe-test: {slug}\n// origin: {origin}\n");
    let body = case.source.trim();
    let program = body.strip_prefix("<?php").unwrap_or(body).trim_start();

    let Some(expected) = case.expected.as_ref() else {
        return Emitted {
            text: format!("<?php\n{header}// vybe-test-mode: compile\n\n{program}\n"),
            pairing: Pairing::Direct,
        };
    };

    // `?>` leaves PHP mode, so appending the check would emit it as HTML; a
    // program running its own `ob_start` would have the buffer taken from
    // underneath it.
    let reason = if body.contains("?>") {
        Some("closes PHP mode (`?>`) — cannot append the check".to_string())
    } else if program.contains("ob_start") {
        Some("uses its own output buffer".to_string())
    } else {
        None
    };
    if let Some(reason) = reason {
        return Emitted {
            text: format!("<?php\n{header}\n{program}\n"),
            pairing: Pairing::Unpairable(reason),
        };
    }

    let want = php_string(&expected.join("\n"));
    Emitted {
        text: format!(
            "<?php\n{header}\n{harness}\n\nob_start();\n\n{program}\n\n__vybe_check(ob_get_clean(), {want});\n"
        ),
        pairing: Pairing::Direct,
    }
}

/// A PHP double-quoted string. `$` must be escaped or PHP interpolates it.
fn php_string(text: &str) -> String {
    let mut out = String::with_capacity(text.len() + 2);
    out.push('"');
    for ch in text.chars() {
        match ch {
            '\\' => out.push_str("\\\\"),
            '"' => out.push_str("\\\""),
            '$' => out.push_str("\\$"),
            '\n' => out.push_str("\\n"),
            '\r' => out.push_str("\\r"),
            '\t' => out.push_str("\\t"),
            _ => out.push(ch),
        }
    }
    out.push('"');
    out
}
