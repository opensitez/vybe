//! Optional owned, typed capture view for incremental walker migration.
//! Only the initial conversion owns a capture allocation; handles clone Rc.
use crate::{CaptureRule, ParseTree, Span, tree::Node};
use std::{marker::PhantomData, rc::Rc};

pub trait RuleIdentity: Copy {
    fn from_capture(capture: CaptureRule) -> Option<Self>;
}

pub struct Pair<'i, R> {
    source: &'i str,
    nodes: Rc<[Node]>,
    index: usize,
    rule: PhantomData<R>,
}
pub struct Pairs<'i, R> {
    source: &'i str,
    nodes: Rc<[Node]>,
    next: usize,
    end: usize,
    rule: PhantomData<R>,
}
impl<R> Clone for Pair<'_, R> {
    fn clone(&self) -> Self {
        Self {
            source: self.source,
            nodes: self.nodes.clone(),
            index: self.index,
            rule: PhantomData,
        }
    }
}
impl<R> Clone for Pairs<'_, R> {
    fn clone(&self) -> Self {
        Self {
            source: self.source,
            nodes: self.nodes.clone(),
            next: self.next,
            end: self.end,
            rule: PhantomData,
        }
    }
}
impl<'i, R: RuleIdentity> Pair<'i, R> {
    pub fn as_rule(&self) -> R {
        R::from_capture(self.nodes[self.index].rule)
            .expect("capture rule missing from generated enum")
    }
    pub fn as_str(&self) -> &'i str {
        let span = self.nodes[self.index].span;
        &self.source[span.start..span.end]
    }
    pub fn as_span(&self) -> SourceSpan<'i> {
        SourceSpan {
            source: self.source,
            span: self.nodes[self.index].span,
        }
    }
    pub fn into_inner(self) -> Pairs<'i, R> {
        Pairs {
            source: self.source,
            next: self.index + 1,
            end: self.nodes[self.index].subtree_end,
            nodes: self.nodes,
            rule: PhantomData,
        }
    }
}
impl<'i, R: RuleIdentity> Iterator for Pairs<'i, R> {
    type Item = Pair<'i, R>;
    fn next(&mut self) -> Option<Self::Item> {
        if self.next >= self.end {
            return None;
        }
        let index = self.next;
        self.next = self.nodes[index].subtree_end;
        Some(Pair {
            source: self.source,
            nodes: self.nodes.clone(),
            index,
            rule: PhantomData,
        })
    }
}
impl<R: RuleIdentity> std::iter::FusedIterator for Pairs<'_, R> {}
impl<'g, 'i> ParseTree<'g, 'i> {
    pub fn into_typed_pairs<R: RuleIdentity>(self) -> Pairs<'i, R> {
        let end = self.nodes.len();
        Pairs {
            source: self.source,
            nodes: self.nodes.into(),
            next: 0,
            end,
            rule: PhantomData,
        }
    }
}

#[derive(Debug, Clone, Copy)]
pub struct SourceSpan<'i> {
    source: &'i str,
    span: Span,
}
impl<'i> SourceSpan<'i> {
    pub fn start(self) -> usize {
        self.span.start
    }
    pub fn end(self) -> usize {
        self.span.end
    }
    pub fn as_str(self) -> &'i str {
        &self.source[self.span.start..self.span.end]
    }
    pub fn start_pos(self) -> Position<'i> {
        Position {
            source: self.source,
            byte: self.span.start,
        }
    }
    pub fn end_pos(self) -> Position<'i> {
        Position {
            source: self.source,
            byte: self.span.end,
        }
    }
}
#[derive(Debug, Clone, Copy)]
pub struct Position<'i> {
    source: &'i str,
    byte: usize,
}
impl Position<'_> {
    pub fn pos(self) -> usize {
        self.byte
    }
    /// Migration convention: one-based scalar columns and LF line boundaries.
    /// This scans the prefix; repeated AST/LSP queries should use source::Index.
    pub fn line_col(self) -> (usize, usize) {
        let prefix = &self.source[..self.byte];
        let line = prefix.bytes().filter(|byte| *byte == b'\n').count() + 1;
        let column = prefix.rsplit('\n').next().unwrap().chars().count() + 1;
        (line, column)
    }
}
