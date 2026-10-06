//! Optional owned, typed capture view for incremental walker migration.
//! Conversion shares the existing capture buffer without copying it; handles clone Rc.
use crate::{
    CaptureRule, ParseTree, Span,
    tree::{Node, Nodes},
};
use std::{marker::PhantomData, rc::Rc};

pub trait RuleIdentity: Copy {
    fn from_capture(capture: CaptureRule) -> Option<Self>;
}

struct TypedContext<'i> {
    source: &'i str,
    nodes: Nodes,
}

// Cache immutable decoded fields once per handle. Sharing the source and arena
// in one context keeps handles smaller than carrying a source slice separately.
pub struct Pair<'i, R> {
    context: Rc<TypedContext<'i>>,
    index: usize,
    node: Node,
    rule: PhantomData<R>,
}
pub struct Pairs<'i, R> {
    context: Rc<TypedContext<'i>>,
    next: usize,
    end: usize,
    rule: PhantomData<R>,
}
impl<R> Clone for Pair<'_, R> {
    fn clone(&self) -> Self {
        Self {
            context: self.context.clone(),
            index: self.index,
            node: self.node,
            rule: PhantomData,
        }
    }
}
impl<R> Clone for Pairs<'_, R> {
    fn clone(&self) -> Self {
        Self {
            context: self.context.clone(),
            next: self.next,
            end: self.end,
            rule: PhantomData,
        }
    }
}
impl<'i, R: RuleIdentity> Pair<'i, R> {
    pub fn as_rule(&self) -> R {
        R::from_capture(CaptureRule::decode(self.node.rule))
            .expect("capture rule missing from generated enum")
    }
    pub fn as_str(&self) -> &'i str {
        let span = self.node.span;
        &self.context.source[span.start..span.end]
    }
    pub fn as_span(&self) -> SourceSpan<'i> {
        SourceSpan {
            source: self.context.source,
            span: self.node.span,
        }
    }
    pub fn into_inner(self) -> Pairs<'i, R> {
        Pairs {
            next: self.index + 1,
            end: self.node.subtree_end,
            context: self.context,
            rule: PhantomData,
        }
    }
}
impl<'i, R: RuleIdentity> Pairs<'i, R> {
    /// Inspect the next sibling without advancing or copying capture storage.
    pub fn peek(&self) -> Option<Pair<'i, R>> {
        (self.next < self.end).then(|| Pair {
            context: self.context.clone(),
            index: self.next,
            node: self.context.nodes.get(self.next),
            rule: PhantomData,
        })
    }
}
impl<'i, R: RuleIdentity> Iterator for Pairs<'i, R> {
    type Item = Pair<'i, R>;
    fn next(&mut self) -> Option<Self::Item> {
        if self.next >= self.end {
            return None;
        }
        let index = self.next;
        let node = self.context.nodes.get(index);
        self.next = node.subtree_end;
        Some(Pair {
            context: self.context.clone(),
            index,
            node,
            rule: PhantomData,
        })
    }
}
impl<R: RuleIdentity> std::iter::FusedIterator for Pairs<'_, R> {}
impl<'g, 'i> ParseTree<'g, 'i> {
    pub fn into_typed_pairs<R: RuleIdentity>(self) -> Pairs<'i, R> {
        let end = self.nodes.len();
        Pairs {
            context: Rc::new(TypedContext {
                source: self.source,
                nodes: self.nodes,
            }),
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

#[cfg(test)]
mod tests {
    use super::*;
    #[derive(Debug, Clone, Copy, PartialEq, Eq)]
    struct Identity(CaptureRule);
    impl RuleIdentity for Identity {
        fn from_capture(capture: CaptureRule) -> Option<Self> {
            Some(Self(capture))
        }
    }
    #[test]
    fn cached_fields_and_shared_lifetimes_preserve_compact_and_wide_views() {
        let grammar = crate::compile(r#"entry = { "a" }"#).unwrap();
        for wide in [false, true] {
            let mut nodes = if wide {
                Nodes::Wide(Vec::new())
            } else {
                Nodes::new(1)
            };
            nodes.push(Node {
                rule: 0,
                span: Span { start: 0, end: 1 },
                subtree_end: 1,
            });
            nodes.push(Node {
                rule: usize::MAX,
                span: Span { start: 1, end: 1 },
                subtree_end: 2,
            });
            nodes.finish(0, 1);
            let tree = ParseTree {
                grammar: &grammar,
                source: "a",
                nodes,
                consumed: 1,
            };
            let mut pairs = tree.into_typed_pairs::<Identity>();
            let root = pairs.peek().unwrap();
            assert_eq!(root.as_rule(), Identity(CaptureRule::User(0)));
            assert_eq!(root.as_str(), "a");
            assert_eq!((root.as_span().start(), root.as_span().end()), (0, 1));
            let cloned = root.clone();
            assert_eq!(cloned.as_rule(), root.as_rule());
            let mut children = cloned.into_inner();
            assert_eq!(pairs.next().unwrap().as_rule(), root.as_rule());
            assert!(pairs.peek().is_none());
            drop(pairs);
            drop(root);
            let eof = children.peek().unwrap();
            assert_eq!(eof.as_rule(), Identity(CaptureRule::EndOfInput));
            assert_eq!(eof.as_str(), "");
            assert_eq!((eof.as_span().start(), eof.as_span().end()), (1, 1));
            assert_eq!(children.next().unwrap().as_rule(), eof.as_rule());
            assert!(children.next().is_none());
        }
    }
}
