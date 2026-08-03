//! Colour, but only for a human looking at a terminal.
//!
//! A report that is read twice — once on screen, once as a log or a diff
//! against `<lang>.tests.txt` — cannot carry escape sequences into the second
//! reading. So the decision is made once, from the stream itself: colour when
//! stdout is a TTY, plain bytes when it is a pipe or a file. Nothing downstream
//! has to strip anything.

use std::io::IsTerminal;
use std::sync::OnceLock;

fn enabled() -> bool {
    static ON: OnceLock<bool> = OnceLock::new();
    // `NO_COLOR` is honoured because a TTY is not consent — a terminal that
    // renders escapes badly, or a CI job with a pty, both set it.
    *ON.get_or_init(|| std::env::var_os("NO_COLOR").is_none() && std::io::stdout().is_terminal())
}

fn paint(code: &str, text: &str) -> String {
    if enabled() { format!("\x1b[{code}m{text}\x1b[0m") } else { text.to_string() }
}

pub fn green(text: &str) -> String {
    paint("32", text)
}

pub fn red(text: &str) -> String {
    paint("31", text)
}

/// A timeout still reports the word `FAILED` — the verdict vocabulary stays
/// cargo's — but wears a different colour, because "hung" and "wrong answer"
/// call for different next steps.
pub fn orange(text: &str) -> String {
    // 256-colour orange, falling back to nothing when colour is off.
    paint("38;5;208", text)
}

pub fn yellow(text: &str) -> String {
    paint("33", text)
}

/// Dim — for the parts of a row that carry no verdict (rules, units, ages).
pub fn grey(text: &str) -> String {
    paint("90", text)
}

/// Bold composes with a colour: the escapes are zero-width, so a name wrapped
/// in both still occupies its padded width and the columns stay aligned.
pub fn bold(text: &str) -> String {
    paint("1", text)
}

/// No colour at all — the finished state. Kept as a function so it can sit
/// beside `green`/`red` wherever a row picks its own paint.
pub fn plain(text: &str) -> String {
    text.to_string()
}
