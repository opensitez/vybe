//! Ruby emitter: one extracted case → a standalone `.rb` test.
//!
//! The simplest emitter here, because it rewrites nothing. `run_ruby_one`
//! joined the program's output lines with `\n` and compared the whole string,
//! so the extracted test does exactly that: the harness collects output, the
//! body is spliced in verbatim, one check at the end.
//!
//! That removes every unpairable category the other emitters have to detect —
//! loops, mismatched print counts, newline-less writes — because nothing is
//! being paired. The only case without assertions is one that never had an
//! expected value (`compile_ok`, 626 of them).

use crate::emit::go::Pairing;
use crate::extract::Case;

pub struct Emitted {
    pub text: String,
    pub pairing: Pairing,
}

pub fn emit(case: &Case, origin: &str, slug: &str, harness: &str) -> Emitted {
    let header = format!("# vybe-test: {slug}\n# origin: {origin}\n");
    let body = case.source.trim_end();

    let Some(expected) = case.expected.as_ref() else {
        return Emitted {
            text: format!("{header}# vybe-test-mode: compile\n\n{body}\n"),
            pairing: Pairing::Direct,
        };
    };

    // The program may define its own `puts` — overriding the harness's and
    // writing to real stdout, which the check would never see.
    if defines_puts(body) {
        return Emitted {
            text: format!("{header}\n{body}\n"),
            pairing: Pairing::Unpairable("defines its own `puts`".into()),
        };
    }

    let want = rb_string(&expected.join("\n"));
    Emitted {
        text: format!("{header}\n{harness}\n\n{body}\n\n__vybe_check({want})\n"),
        pairing: Pairing::Direct,
    }
}

fn defines_puts(src: &str) -> bool {
    src.contains("def puts") || src.contains("def self.puts")
}

/// A Ruby double-quoted string. `#` is escaped unconditionally: only `#{`
/// interpolates, but escaping every `#` is correct and cheaper than tracking
/// which ones are followed by a brace.
fn rb_string(text: &str) -> String {
    let mut out = String::with_capacity(text.len() + 2);
    out.push('"');
    for ch in text.chars() {
        match ch {
            '\\' => out.push_str("\\\\"),
            '"' => out.push_str("\\\""),
            '#' => out.push_str("\\#"),
            '\n' => out.push_str("\\n"),
            '\r' => out.push_str("\\r"),
            '\t' => out.push_str("\\t"),
            _ => out.push(ch),
        }
    }
    out.push('"');
    out
}
