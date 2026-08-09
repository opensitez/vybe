//! C emitter: one extracted case → a standalone `.c` test.
//!
//! The program is assembled the way the corpus's own `program_src` does —
//! includes, then declarations, then `int main() { body }` — with the includes
//! and declarations arriving together in `case.prelude`.
//!
//! **Checks are inline and per print.** Every way of accumulating output fails
//! under Vybe while working under `cc`: a file-scope `static` written from a
//! function, `snprintf(buf + n, …)` pointer arithmetic, `strcat`, and varargs
//! `vsnprintf`. So each `printf` is formatted into a fresh LOCAL buffer and
//! compared where it stands. See `harness/c/check.c`.
//!
//! **Failure is `assert(0)`** — the only signal that reaches a non-zero status
//! in both runtimes. `exit(1)` and `return n` both give 0 under Vybe, because
//! `exit_call_is_return = true` compiles `exit` to a RETURN and main's return
//! value is not mapped to the process status.

use crate::emit::go::Pairing;
use crate::extract::Case;

pub struct Emitted {
    pub text: String,
    pub pairing: Pairing,
}

pub fn emit(case: &Case, origin: &str, slug: &str, _harness: &str) -> Emitted {
    let header = format!("// vybe-test: {slug}\n// origin: {origin}\n");
    let body = case.source.trim_end();

    // `c_compile_cases!` asserted `compile_ok` and nothing else. Compile mode
    // runs `vybex -d`, which is that assertion exactly: the frontend accepts
    // the program, and the body never executes — so a case whose body calls
    // `popen`, `mkstemp` or `fork` stays as inert as it was under cargo.
    //
    // Assembling it as a RUN test instead would be worse than useless: main's
    // return value is not the process status under Vybe, so `return EPERM;`
    // would pass whatever the body did, while `cc` would report exit 1 as a
    // failure the corpus never asserted.
    let Some(expected) = case.expected.as_ref() else {
        return Emitted {
            text: format!(
                "{header}// vybe-test-mode: compile\n{}\n",
                assemble(body, case.prelude.as_deref(), false)
            ),
            pairing: Pairing::Direct,
        };
    };

    let prints = find_prints(body);
    if let Some(reason) = unpairable(body, case.prelude.as_deref(), &prints, expected.len()) {
        return Emitted {
            text: format!(
                "{header}{}\n",
                assemble(body, case.prelude.as_deref(), true)
            ),
            pairing: Pairing::Unpairable(reason),
        };
    }

    // RUNTIME pairing: a table of expected lines plus a counter, checked
    // where each print stands. The counter advances at run time, so a print
    // inside a LOOP pairs correctly — which per-print pairing cannot express
    // and output collection cannot achieve either, because every accumulation
    // path is broken under Vybe (see the module docs).
    let mut out = body.to_string();
    for (i, p) in prints.iter().enumerate().rev() {
        let _ = i;
        out.replace_range(p.start..p.end, &check_for(p));
    }

    let table = expected
        .iter()
        .enumerate()
        .map(|(i, w)| {
            // Fall back to the FIRST print when there is no print at index i.
            // One print inside a LOOP produces many lines, so indexing prints
            // by the expected-line index left every entry after the first
            // without its newline — the table read `{"0\n", "1", "2"}`.
            let nl = prints
                .get(i)
                .or_else(|| prints.first())
                .is_some_and(|p| p.fmt_ends_with_newline);
            c_string(&if nl { format!("{w}\n") } else { w.clone() })
        })
        .collect::<Vec<_>>()
        .join(", ");
    // Too FEW lines is a failure too, and only the end can see it.
    let epilogue =
        "if (__i != __n) { printf(\"FAIL: %d line(s), wanted %d\\n\", __i, __n); assert(0); }\n";
    // The prologue declares `__w`/`__n`/`__i`, and every check reads them, so
    // it has to be in scope at every check.
    //
    // Inside `main` is the default. But a case can print from a HELPER defined
    // above main — `typedef void F(void); void f(void) { printf("V"); }` — and
    // then the check lands in `f` while the declarations sit in main's body,
    // which `cc` rejects outright ("undeclared identifier `__i`"). 86 files
    // were emitted that way and could never compile in either runtime.
    //
    // FILE scope fixes those, and is measured to work under BOTH Vybe and cc:
    // a `static` expected table plus a `static` counter incremented from a
    // function. That is not the same thing as the accumulation this module's
    // header rules out — that was a file-scope char BUFFER written from a
    // function, which is broken; a counter and a const table are not.
    //
    // It is used only where needed, so the 6,000-odd cases that already pair
    // keep the placement they were verified with.
    let main_body_at = if declares_main(&out) {
        find_main(&out).and_then(|m| out[m..].find('{').map(|o| m + o + 1))
    } else {
        None
    };
    let checks_outside_main = main_body_at.is_some_and(|at| out[..at].contains("__i++"));
    let (pro_at, file_scope) = match main_body_at {
        Some(_) if checks_outside_main => (0, true),
        Some(at) => (at, false),
        None => (0, false),
    };
    let storage = if file_scope { "static " } else { "" };
    let prologue = format!(
        "{storage}const char *__w[] = {{{table}}};\n{storage}int __n = {}, __i = 0;\n",
        expected.len()
    );
    let epi_at = out.rfind("return 0;").filter(|at| *at >= pro_at);
    let body_out = match epi_at {
        Some(at) => format!(
            "{}{prologue}{}{epilogue}{}",
            &out[..pro_at],
            &out[pro_at..at],
            &out[at..]
        ),
        None => format!("{}{prologue}{}\n{epilogue}", &out[..pro_at], &out[pro_at..]),
    };

    Emitted {
        text: format!(
            "{header}{}\n",
            assemble(&body_out, case.prelude.as_deref(), true)
        ),
        pairing: Pairing::Direct,
    }
}

/// What the corpus's `program_src` built, plus — when `checks` — the headers
/// the check code needs.
fn assemble(body: &str, prelude: Option<&str>, checks: bool) -> String {
    let head = prelude.map(str::trim_end).unwrap_or("");
    let mut out = String::with_capacity(body.len() + head.len() + 128);
    // Headers FIRST: a declaration may itself call `printf`, and an include
    // that follows it would be too late.
    // `string.h` for `strcmp`, `assert.h` for the failure signal, `stdio.h`
    // for `snprintf` — added only when the source has not already.
    //
    // A compile-only case has no check code, and adding them there would break
    // what it measures: `cover_headers_misc`, `cover_wchar_h` and `cover_uchar_h`
    // exist to prove a program compiles with the includes it NAMES. Handing it
    // three more can hide a missing header or invent one.
    for (needle, line) in [
        ("<stdio.h>", "#include <stdio.h>"),
        ("<string.h>", "#include <string.h>"),
        ("<assert.h>", "#include <assert.h>"),
    ] {
        if checks && !head.contains(needle) {
            out.push_str(line);
            out.push('\n');
        }
    }
    out.push_str(head);
    if !head.is_empty() {
        out.push('\n');
    }
    // Some cases carry a COMPLETE program in `body` rather than a main body.
    // Wrapping one again nests `int main` inside `int main`, which cc rejects
    // outright ("function definition is not allowed here").
    if declares_main(body) {
        out.push_str(body.trim_end());
        out.push('\n');
    } else {
        out.push_str("int main() {\n");
        out.push_str(body.trim_end());
        out.push('\n');
        out.push_str("}\n");
    }
    out
}

/// Where `int main` / `void main` actually appears, skipping comments and
/// string literals. A case body can open with a COMMENTED-OUT `int main() {`
/// explaining what is not legal C — matching that put the prologue inside the
/// comment and left the real `main` nested inside the wrapper.
fn find_main(src: &str) -> Option<usize> {
    let bytes = src.as_bytes();
    let mut i = 0usize;
    while i < bytes.len() {
        if let Some(next) = skip_literal(bytes, i) {
            i = next;
            continue;
        }
        if bytes[i] == b'/' {
            match bytes.get(i + 1) {
                Some(b'/') => {
                    i = src[i..]
                        .find('\n')
                        .map(|o| i + o + 1)
                        .unwrap_or(bytes.len());
                    continue;
                }
                Some(b'*') => {
                    i = src[i + 2..]
                        .find("*/")
                        .map(|o| i + 2 + o + 2)
                        .unwrap_or(bytes.len());
                    continue;
                }
                _ => {}
            }
        }
        if src.is_char_boundary(i)
            && (src[i..].starts_with("int main") || src[i..].starts_with("void main"))
        {
            return Some(i);
        }
        i += 1;
    }
    None
}

fn declares_main(src: &str) -> bool {
    find_main(src).is_some()
}

/// The inline check replacing one print. Braced so repeated checks in the same
/// function do not collide on `__t`.
fn check_for(p: &Print) -> String {
    format!(
        "{{ char __t[512]; snprintf(__t, sizeof(__t), {});\n  \
         if (__i >= __n || strcmp(__t, __w[__i]) != 0) {{ \
         printf(\"FAIL at line %d: got [%s]\\n\", __i, __t); assert(0); }} __i++; }}",
        p.args
    )
}

/// Does the format end with `\n`? Only a literal can be read; a format held in
/// a variable is left for `unpairable` to reject.
///
/// The format is not always ONE literal. C concatenates adjacent literals, and
/// `<inttypes.h>` is built on it: `printf("%" PRIx8 "\n", v)` is three tokens
/// and the newline is in the LAST one. Reading only the first (`"%"`) said "no
/// newline", so the expected line was stored without one while the program
/// printed one — every test in `inttypes_format_macro_output` failed on the
/// trailing byte, under `cc` as well as under Vybe.
fn fmt_ends_with_newline(args: &str) -> Option<bool> {
    let head = format_segment(args);
    let bytes = head.as_bytes();
    let mut last: Option<&str> = None;
    let mut i = 0usize;
    while i < bytes.len() {
        match skip_literal(bytes, i) {
            Some(end) => {
                last = Some(&head[i..end]);
                i = end;
            }
            None => i += 1,
        }
    }
    Some(last?.trim_end_matches('"').ends_with("\\n"))
}

/// The argument list up to the first top-level comma — i.e. the format, however
/// many literals and macro names it is spelled with.
fn format_segment(args: &str) -> &str {
    let bytes = args.as_bytes();
    let mut depth = 0i32;
    let mut i = 0usize;
    while i < bytes.len() {
        if let Some(end) = skip_literal(bytes, i) {
            i = end;
            continue;
        }
        match bytes[i] {
            b'(' | b'[' => depth += 1,
            b')' | b']' => depth -= 1,
            b',' if depth == 0 => return &args[..i],
            _ => {}
        }
        i += 1;
    }
    args
}

struct Print {
    start: usize,
    end: usize,
    /// The whole argument list, ready to hand to `snprintf`.
    args: String,
    /// Whether the format literal ends in a newline — decides whether the
    /// expected line gets one appended.
    fmt_ends_with_newline: bool,
}

/// `printf(...)` and `puts(...)` statements. `puts` appends a newline, so it
/// becomes a `"%s\n"` format.
fn find_prints(src: &str) -> Vec<Print> {
    let bytes = src.as_bytes();
    let mut out = Vec::new();
    let mut i = 0usize;
    while i < bytes.len() {
        if let Some(next) = skip_literal(bytes, i) {
            i = next;
            continue;
        }
        let hit = ["printf(", "puts("]
            .into_iter()
            .find(|n| src.is_char_boundary(i) && src[i..].starts_with(n));
        if let Some(needle) = hit {
            let before = if i == 0 { b' ' } else { bytes[i - 1] };
            // Not `sprintf(` / `fprintf(` / a member access.
            if !is_ident(before) {
                let open = i + needle.len();
                if let Some(close) = close_paren(bytes, open) {
                    let inner = src[open..close].trim().to_string();
                    let args = if needle == "puts(" {
                        format!("\"%s\\n\", {inner}")
                    } else {
                        inner
                    };
                    // Swallow a trailing `;`.
                    let mut end = close + 1;
                    if bytes.get(end) == Some(&b';') {
                        end += 1;
                    }
                    let Some(nl) = fmt_ends_with_newline(&args) else {
                        i = end;
                        continue;
                    };
                    out.push(Print {
                        start: i,
                        end,
                        args,
                        fmt_ends_with_newline: nl,
                    });
                    i = end;
                    continue;
                }
            }
        }
        i += 1;
    }
    out
}

/// Is `at` inside a `#define`'s logical line? A trailing `\` continues it onto
/// the next physical line, so the run does not simply end at the first newline.
fn in_define(src: &str, at: usize) -> bool {
    let mut pos = 0usize;
    let mut in_macro = false;
    for line in src.split_inclusive('\n') {
        let end = pos + line.len();
        if !in_macro {
            in_macro = line.trim_start().starts_with("#define");
        }
        if in_macro {
            if at >= pos && at < end {
                return true;
            }
            // The continuation only holds while the line ends in a backslash.
            in_macro = line.trim_end().ends_with('\\');
        }
        pos = end;
    }
    false
}

fn unpairable(
    src: &str,
    prelude: Option<&str>,
    prints: &[Print],
    expected: usize,
) -> Option<String> {
    if prints.is_empty() {
        return Some("no printf/puts to pair".into());
    }
    // Declarations can print too, and those calls are not in `body`.
    if prelude.is_some_and(|p| p.contains("printf(") || p.contains("puts(")) {
        return Some("prints from a declaration — not reachable from the body".into());
    }
    // A `#define` body is ONE logical line. The check spans several, so
    // substituting it there truncates the macro at the first newline and `cc`
    // rejects the file outright ("expected identifier or '('").
    if prints.iter().any(|p| in_define(src, p.start)) {
        return Some("print inside a #define — the check cannot span the macro's line".into());
    }
    if src.contains("fprintf(") || src.contains("putchar(") || src.contains("fwrite(") {
        return Some("fprintf/putchar/fwrite — output is not one line per call".into());
    }
    // A count mismatch is no longer a reason to give up: the counter checks it
    // at run time, which also covers a loop whose trip count is not static.
    let _ = expected;
    None
}

fn close_paren(bytes: &[u8], from: usize) -> Option<usize> {
    let mut depth = 1usize;
    let mut i = from;
    while i < bytes.len() {
        if let Some(next) = skip_literal(bytes, i) {
            i = next;
            continue;
        }
        match bytes[i] {
            b'(' => depth += 1,
            b')' => {
                depth -= 1;
                if depth == 0 {
                    return Some(i);
                }
            }
            _ => {}
        }
        i += 1;
    }
    None
}

fn skip_literal(bytes: &[u8], at: usize) -> Option<usize> {
    let quote = match bytes.get(at)? {
        b'"' => b'"',
        b'\'' => b'\'',
        _ => return None,
    };
    let mut i = at + 1;
    while i < bytes.len() {
        match bytes[i] {
            b'\\' => i += 2,
            b if b == quote => return Some(i + 1),
            _ => i += 1,
        }
    }
    Some(bytes.len())
}

fn is_ident(byte: u8) -> bool {
    byte.is_ascii_alphanumeric() || byte == b'_'
}

fn c_string(text: &str) -> String {
    let mut out = String::with_capacity(text.len() + 2);
    out.push('"');
    for ch in text.chars() {
        match ch {
            '\\' => out.push_str("\\\\"),
            '"' => out.push_str("\\\""),
            '\n' => out.push_str("\\n"),
            '\r' => out.push_str("\\r"),
            '\t' => out.push_str("\\t"),
            _ => out.push(ch),
        }
    }
    out.push('"');
    out
}
