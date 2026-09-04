//! Byte offset to `(line, column)` for the source a front end is walking.
//!
//! ⛔⛔ THIS EXISTS BECAUSE WALKERS WERE QUADRATIC IN PROGRAM SIZE. A walker's
//! `to_span` asks pest's `Position::line_col`, which counts newlines FROM THE
//! START OF THE INPUT — O(offset) — and it asks twice for every node, so the
//! total is O(nodes x length). Measured on VB with `-c` over a ladder of `Dim`
//! lines: 320 -> 2.89s, 640 -> 10.69s, 1280 -> 42.60s, i.e. 4x the time per
//! doubling where linear is 2x, and `line_col` was the hottest frame in the
//! profile. With the index: 0.85s / 1.42s / 3.01s.
//!
//! The index is a thread-local, so `to_span` stays a free function and no
//! walker has to thread state through its call graph.
//!
//! ⛔ This module takes BYTE OFFSETS, not a pest `Pair`. `vybe_ast` has no
//! dependencies and keeps none — the caller pulls the offsets off its own
//! `Span` type and applies its own line/column base.

use crate::Span;
use std::cell::RefCell;

thread_local! {
    /// Byte offset of the start of every line in the source being walked on
    /// this thread. Empty when no parse has installed one.
    static LINE_STARTS: RefCell<Vec<usize>> = const { RefCell::new(Vec::new()) };
}

/// The line index for one parse, installed on construction and replaced by
/// whatever was there before when it drops.
///
/// ⛔ THE RESTORE IS THE POINT. A walker that parses a second source mid-walk —
/// an expression fragment, an embedded literal, a synthesized helper — installs
/// an index for THAT text. Without the restore every node walked afterwards is
/// numbered against the fragment.
///
/// ```ignore
/// let _index = LineIndex::install(source);
/// let pairs = MyParser::parse(Rule::program, source)?;
/// ```
pub struct LineIndex(Vec<usize>);

impl LineIndex {
    /// Index `src` and install it, handing back a guard that restores the
    /// previous index.
    pub fn install(src: &str) -> Self {
        let mut starts = Vec::with_capacity(src.len() / 24 + 1);
        starts.push(0);
        starts.extend(src.match_indices('\n').map(|(i, _)| i + 1));
        LineIndex(LINE_STARTS.with(|c| c.replace(starts)))
    }
}

impl Drop for LineIndex {
    fn drop(&mut self) {
        LINE_STARTS.with(|c| *c.borrow_mut() = std::mem::take(&mut self.0));
    }
}

/// 1-based `(line, column)` for a byte offset, matching pest's `line_col`, or
/// `None` when no index is installed.
///
/// A caller that gets `None` has no index for this parse and must fall back to
/// whatever it did before — answering a zero span instead would silently strip
/// the line numbers out of every diagnostic.
pub fn line_col(offset: usize) -> Option<(u32, u32)> {
    LINE_STARTS.with(|c| {
        let starts = c.borrow();
        if starts.is_empty() {
            return None;
        }
        // The last line start at or before `offset`.
        let idx = starts.partition_point(|&s| s <= offset) - 1;
        Some((idx as u32 + 1, (offset - starts[idx]) as u32 + 1))
    })
}

/// Both ends of a span in ONE thread-local lookup.
///
/// ⛔ A span costs one `with`, not two. Calling [`line_col`] per end doubles
/// the thread-local access and the borrow, and this crate is compiled
/// separately from its callers, so nothing inlines those away: measured on the
/// VB 1280-line ladder, two lookups per span cost 5.04s against 3.01s for one.
fn ends(start: usize, end: usize) -> Option<((u32, u32), (u32, u32))> {
    LINE_STARTS.with(|c| {
        let starts = c.borrow();
        if starts.is_empty() {
            return None;
        }
        let at = |offset: usize| {
            let idx = starts.partition_point(|&s| s <= offset) - 1;
            (idx as u32 + 1, (offset - starts[idx]) as u32 + 1)
        };
        Some((at(start), at(end)))
    })
}

/// A span whose lines and columns both count from 1, as pest reports them.
pub fn span_1based(start: usize, end: usize) -> Option<Span> {
    let ((start_line, start_col), (end_line, end_col)) = ends(start, end)?;
    Some(Span {
        start_line,
        start_col,
        end_line,
        end_col,
    })
}

/// A span whose lines and columns both count from 0.
pub fn span_0based(start: usize, end: usize) -> Option<Span> {
    let ((start_line, start_col), (end_line, end_col)) = ends(start, end)?;
    Some(Span {
        start_line: start_line - 1,
        start_col: start_col - 1,
        end_line: end_line - 1,
        end_col: end_col - 1,
    })
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn reports_pest_line_col_for_every_offset() {
        let src = "alpha\nbeta\n\ngamma";
        let _index = LineIndex::install(src);
        assert_eq!(line_col(0), Some((1, 1)));
        assert_eq!(line_col(5), Some((1, 6))); // the newline itself
        assert_eq!(line_col(6), Some((2, 1)));
        assert_eq!(line_col(11), Some((3, 1))); // the empty line
        assert_eq!(line_col(12), Some((4, 1)));
        assert_eq!(
            span_0based(6, 10),
            Some(Span {
                start_line: 1,
                start_col: 0,
                end_line: 1,
                end_col: 4
            })
        );
    }

    #[test]
    fn answers_none_without_an_index() {
        // Proves the check can fail: a walker that never installs an index
        // keeps its own fallback rather than reporting line 0.
        assert_eq!(line_col(0), None);
        assert_eq!(span_1based(0, 1), None);
    }

    #[test]
    fn a_nested_parse_restores_the_outer_index() {
        let outer = "one\ntwo\nthree";
        let _index = LineIndex::install(outer);
        assert_eq!(line_col(8), Some((3, 1)));
        {
            let _fragment = LineIndex::install("x");
            assert_eq!(line_col(0), Some((1, 1)));
        }
        assert_eq!(line_col(8), Some((3, 1)));
    }
}
