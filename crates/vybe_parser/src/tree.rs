//! A contiguous preorder capture arena, independent of pest's representation.
//! Cheap borrowed handles expose a migration-friendly view over that arena.
use crate::grammar::RuleId;
use crate::{Span, program::Program};

#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum CaptureRule {
    User(RuleId),
    EndOfInput,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub(crate) struct Node {
    pub rule: CaptureRule,
    pub span: Span,
    pub subtree_end: usize,
}

#[derive(Debug)]
pub struct ParseTree<'g, 'i> {
    pub(crate) grammar: &'g dyn Program,
    pub(crate) source: &'i str,
    pub(crate) nodes: Vec<Node>,
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
        self.tree.nodes[self.index].rule
    }
    pub fn rule_name(&self) -> &'g str {
        match self.as_rule() {
            CaptureRule::User(rule) => self.tree.grammar.rule(rule).name,
            CaptureRule::EndOfInput => "EOI",
        }
    }
    pub fn as_span(&self) -> Span {
        self.tree.nodes[self.index].span
    }
    pub fn as_str(&self) -> &'i str {
        let span = self.as_span();
        &self.tree.source[span.start..span.end]
    }
    pub fn into_inner(self) -> Pairs<'t, 'g, 'i> {
        Pairs {
            tree: self.tree,
            next: self.index + 1,
            end: self.tree.nodes[self.index].subtree_end,
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
        self.next = self.tree.nodes[index].subtree_end;
        Some(Pair {
            tree: self.tree,
            index,
        })
    }
}

impl std::iter::FusedIterator for Pairs<'_, '_, '_> {}
