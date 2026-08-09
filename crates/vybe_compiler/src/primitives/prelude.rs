//! Prelude machinery shared by every language that carries one.
//!
//! A prelude is source in the language under compilation that is prepended to
//! the program body — exception classes, helper functions, singletons. Four
//! languages carry one, and each had independently written the SAME two
//! things:
//!
//! - a content-keyed cache of parsed statements, so fixed prelude source is
//!   walked once per process instead of once per compile —
//!   `go/walker.rs:281`, `python/walker.rs:498`, `php/walker.rs:3982`,
//!   `js/lib.rs:130`;
//! - the splice that puts a group ahead of the body.
//!
//! Four copies meant four places to change. The change we want next is a
//! bigger one: the parse is already cached, but the prelude is RE-COMPILED on
//! every program, and that is where the time actually goes (measured: a php
//! program that triggers one ~10-class group costs 5x one that triggers none).
//! Caching compiled output belongs behind this seam, written once.
//!
//! ## What is NOT here
//!
//! WHICH groups a program needs. That decision is language semantics — php
//! inspects its parsed statements, python matches its source text, go uses
//! flags collected during the walk — and moving it here would put per-language
//! knowledge into shared code. The language decides; this module remembers.

use vybe_ast::Statement;

/// Most a single process will retain. Prelude source is normally a handful of
/// fixed constants, but a language may build one per program (python does),
/// and an unbounded map in a long-lived worker is a leak rather than a cache.
const MAX_ENTRIES: usize = 256;

/// Parse `src` once per process and hand back an owned copy.
///
/// Keyed by CONTENT, not by pointer or by a group tag: a dynamically-built
/// prelude is then correct rather than aliased onto whatever was cached under
/// the same name. Callers mutate what they get (splicing is `append`), so each
/// receives its own clone — cloning statements is far cheaper than re-walking
/// the source, which is the whole point.
///
/// `parse` runs only on a miss. It is the LANGUAGE's parser: preludes are
/// written in the language under compilation, so there is no shared parse to
/// call here.
pub fn cached<E>(
    src: &str,
    parse: impl FnOnce(&str) -> Result<Vec<Statement>, E>,
) -> Result<Vec<Statement>, E> {
    use std::collections::HashMap;
    use std::sync::{Mutex, OnceLock};
    static CACHE: OnceLock<Mutex<HashMap<String, Vec<Statement>>>> = OnceLock::new();
    let cache = CACHE.get_or_init(|| Mutex::new(HashMap::new()));

    if let Some(hit) = cache.lock().unwrap().get(src) {
        return Ok(hit.clone());
    }
    let parsed = parse(src)?;
    let mut guard = cache.lock().unwrap();
    // Re-check: two threads can miss on the same source concurrently, and the
    // second must not push the map past the bound for an entry already in it.
    if !guard.contains_key(src) && guard.len() < MAX_ENTRIES {
        guard.insert(src.to_string(), parsed.clone());
    }
    Ok(parsed)
}

/// [`cached`] for a parser that cannot fail.
pub fn cached_infallible(src: &str, parse: impl FnOnce(&str) -> Vec<Statement>) -> Vec<Statement> {
    let out: Result<Vec<Statement>, std::convert::Infallible> = cached(src, |s| Ok(parse(s)));
    out.unwrap_or_default()
}

/// One declared prelude group: some source, and whether THIS program needs it.
///
/// The language fills `needed` however it likes — php inspects parsed
/// statements, python matches source text, go reads flags collected during the
/// walk. All three answer the same question, so only the answer crosses into
/// shared code.
pub struct Group<'a> {
    /// Identifies the group in diagnostics. Not a dispatch key: nothing here
    /// branches on it, so no per-language name ever reaches shared logic.
    pub name: &'static str,
    pub source: &'a str,
    pub needed: bool,
}

impl<'a> Group<'a> {
    pub fn new(name: &'static str, source: &'a str, needed: bool) -> Self {
        Group {
            name,
            source,
            needed,
        }
    }
}

/// Parse the needed groups (cached) and return them spliced together in
/// DECLARATION order, ready to sit ahead of the program body.
///
/// Declaration order is the contract: a group that defines what a later group
/// references has to come first, and the language expresses that by the order
/// it lists them — not by remembering to splice in reverse at each call site,
/// which is what the hand-rolled version required.
///
/// Groups whose `needed` is false are never parsed, so an unused group costs
/// nothing beyond the flag that decided it.
pub fn build<E>(
    groups: &[Group<'_>],
    mut parse: impl FnMut(&str) -> Result<Vec<Statement>, E>,
) -> Result<Vec<Statement>, E> {
    let mut out: Vec<Statement> = Vec::new();
    for group in groups.iter().filter(|g| g.needed) {
        out.append(&mut cached(group.source, &mut parse)?);
    }
    Ok(out)
}

/// [`build`] for a parser that cannot fail.
pub fn build_infallible(
    groups: &[Group<'_>],
    mut parse: impl FnMut(&str) -> Vec<Statement>,
) -> Vec<Statement> {
    let out: Result<Vec<Statement>, std::convert::Infallible> = build(groups, |s| Ok(parse(s)));
    out.unwrap_or_default()
}

/// Splice `group` ahead of `body`, preserving the order of earlier splices.
///
/// The idiom this replaces was written out at every call site as
/// `let mut p = parse(G); p.append(&mut body); body = p;` — correct, but it
/// reads as a rebuild rather than an insert, and one site getting the operand
/// order wrong silently puts the prelude AFTER the code that needs it.
pub fn prepend(body: &mut Vec<Statement>, mut group: Vec<Statement>) {
    if group.is_empty() {
        return;
    }
    group.append(body);
    *body = group;
}
