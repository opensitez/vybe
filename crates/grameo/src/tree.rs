//! A contiguous preorder capture arena, independent of pest's representation.
//! Cheap borrowed handles expose a migration-friendly view over that arena.
use crate::grammar::RuleId;
use crate::{Span, program::Program};

#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum CaptureRule {
    User(RuleId),
    EndOfInput,
}

impl CaptureRule {
    pub(crate) fn encode(self) -> usize {
        match self {
            Self::User(rule) => rule,
            Self::EndOfInput => usize::MAX,
        }
    }
    pub(crate) fn decode(rule: usize) -> Self {
        if rule == usize::MAX {
            Self::EndOfInput
        } else {
            Self::User(rule)
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) struct Node {
    pub rule: usize,
    pub span: Span,
    pub subtree_end: usize,
}

#[derive(Debug, Clone, Copy)]
pub(crate) struct CompactNode {
    rule: u32,
    start: u32,
    end: u32,
    subtree_end: u32,
}
impl CompactNode {
    fn expand(self) -> Node {
        Node {
            rule: if self.rule == u32::MAX {
                usize::MAX
            } else {
                self.rule as usize
            },
            span: Span {
                start: self.start as usize,
                end: self.end as usize,
            },
            subtree_end: self.subtree_end as usize,
        }
    }
}
#[derive(Debug)]
pub(crate) enum Nodes {
    Compact(Vec<CompactNode>),
    Wide(Vec<Node>),
}
impl Nodes {
    pub(crate) fn new(source_len: usize) -> Self {
        if source_len <= u32::MAX as usize {
            Self::Compact(Vec::new())
        } else {
            Self::Wide(Vec::new())
        }
    }
    pub(crate) fn reserve(&mut self, additional: usize) {
        match self {
            Self::Compact(nodes) => nodes.reserve(additional),
            Self::Wide(nodes) => nodes.reserve(additional),
        }
    }
    #[inline]
    pub(crate) fn len(&self) -> usize {
        match self {
            Self::Compact(v) => v.len(),
            Self::Wide(v) => v.len(),
        }
    }
    pub(crate) fn iter(&self) -> impl Iterator<Item = Node> {
        (0..self.len()).map(|index| self.get(index))
    }
    #[inline]
    pub(crate) fn get(&self, index: usize) -> Node {
        match self {
            Self::Compact(v) => v[index].expand(),
            Self::Wide(v) => v[index],
        }
    }
    #[inline]
    pub(crate) fn truncate(&mut self, len: usize) {
        match self {
            Self::Compact(v) => v.truncate(len),
            Self::Wide(v) => v.truncate(len),
        }
    }
    #[inline]
    pub(crate) fn finish(&mut self, index: usize, end: usize) {
        let len = self.len();
        match self {
            Self::Compact(v) => {
                debug_assert!(end <= u32::MAX as usize);
                v[index].end = end as u32;
                v[index].subtree_end = len as u32;
            }
            Self::Wide(v) => {
                v[index].span.end = end;
                v[index].subtree_end = len;
            }
        }
    }
    #[inline]
    pub(crate) fn push(&mut self, node: Node) {
        if let Self::Compact(v) = self {
            if (node.rule < u32::MAX as usize || node.rule == usize::MAX)
                && node.subtree_end <= u32::MAX as usize
            {
                debug_assert!(
                    node.span.start <= u32::MAX as usize && node.span.end <= u32::MAX as usize
                );
                v.push(CompactNode {
                    rule: node.rule as u32,
                    start: node.span.start as u32,
                    end: node.span.end as u32,
                    subtree_end: node.subtree_end as u32,
                });
                return;
            }
            *self = Self::Wide(v.iter().copied().map(CompactNode::expand).collect());
        }
        let Self::Wide(v) = self else { unreachable!() };
        v.push(node);
    }
}

#[derive(Debug)]
pub struct ParseTree<'g, 'i> {
    pub(crate) grammar: &'g dyn Program,
    pub(crate) source: &'i str,
    pub(crate) nodes: Nodes,
    pub(crate) consumed: usize,
}

impl<'g, 'i> ParseTree<'g, 'i> {
    pub fn consumed(&self) -> usize {
        self.consumed
    }
    pub fn source(&self) -> &'i str {
        self.source
    }
    pub fn pairs(&self) -> Pairs<'_, 'g, 'i> {
        Pairs {
            tree: self,
            next: 0,
            end: self.nodes.len(),
        }
    }
    pub fn capture_count(&self) -> usize {
        self.nodes.len()
    }
}

#[derive(Debug, Clone, Copy)]
pub struct Pair<'t, 'g, 'i> {
    tree: &'t ParseTree<'g, 'i>,
    index: usize,
}

impl<'t, 'g, 'i> Pair<'t, 'g, 'i> {
    pub fn as_rule(&self) -> CaptureRule {
        CaptureRule::decode(self.tree.nodes.get(self.index).rule)
    }
    pub fn rule_name(&self) -> &'g str {
        match self.as_rule() {
            CaptureRule::User(rule) => self.tree.grammar.rule(rule).name,
            CaptureRule::EndOfInput => "EOI",
        }
    }
    pub fn as_span(&self) -> Span {
        self.tree.nodes.get(self.index).span
    }
    pub fn as_str(&self) -> &'i str {
        let span = self.as_span();
        &self.tree.source[span.start..span.end]
    }
    pub fn into_inner(self) -> Pairs<'t, 'g, 'i> {
        Pairs {
            tree: self.tree,
            next: self.index + 1,
            end: self.tree.nodes.get(self.index).subtree_end,
        }
    }
}

#[derive(Debug, Clone)]
pub struct Pairs<'t, 'g, 'i> {
    tree: &'t ParseTree<'g, 'i>,
    next: usize,
    end: usize,
}

impl<'t, 'g, 'i> Iterator for Pairs<'t, 'g, 'i> {
    type Item = Pair<'t, 'g, 'i>;
    fn next(&mut self) -> Option<Self::Item> {
        if self.next >= self.end {
            return None;
        }
        let index = self.next;
        self.next = self.tree.nodes.get(index).subtree_end;
        Some(Pair {
            tree: self.tree,
            index,
        })
    }
}

impl std::iter::FusedIterator for Pairs<'_, '_, '_> {}

#[cfg(test)]
mod compact_tests {
    use super::*;
    #[test]
    fn arena_promotes_large_rules_without_changing_existing_eof_or_spans() {
        if usize::BITS <= 32 {
            return;
        }
        let mut nodes = Nodes::new(100);
        let eof = Node {
            rule: usize::MAX,
            span: Span { start: 4, end: 4 },
            subtree_end: 1,
        };
        nodes.push(eof);
        let large = Node {
            rule: u32::MAX as usize,
            span: Span { start: 4, end: 4 },
            subtree_end: 2,
        };
        nodes.push(large);
        assert!(matches!(nodes, Nodes::Wide(_)));
        assert_eq!(nodes.get(0), eof);
        assert_eq!(nodes.get(1), large);
        nodes.truncate(1);
        assert_eq!(nodes.get(0), eof);
        nodes.finish(0, 8);
        assert_eq!(nodes.get(0).span, Span { start: 4, end: 8 });
    }
    #[test]
    fn wide_source_spans_and_zero_width_children_preserve_preorder() {
        let Some(end) = (u32::MAX as usize).checked_add(9) else {
            return;
        };
        let mut nodes = Nodes::new(end);
        let node = Node {
            rule: 1,
            span: Span {
                start: end - 3,
                end,
            },
            subtree_end: 1,
        };
        nodes.push(node);
        assert_eq!(nodes.get(0), node);
        let mut nodes = Nodes::new(9);
        nodes.push(Node {
            rule: 0,
            span: Span { start: 2, end: 0 },
            subtree_end: 1,
        });
        nodes.push(Node {
            rule: usize::MAX,
            span: Span { start: 2, end: 2 },
            subtree_end: 2,
        });
        nodes.finish(0, 2);
        assert_eq!(nodes.get(0).subtree_end, 2);
        assert_eq!(nodes.get(1).rule, usize::MAX);
    }
}
