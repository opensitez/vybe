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
    /// Resolve a finish-hook token under this anchor without scanning captures.
    /// No handle is produced for another input, stale/pruned node, wrong rule
    /// or span, or a token outside the anchor's subtree.
    pub fn resolve_capture(&self, handle: crate::tree::CaptureHandle) -> Option<Self> {
        if handle.source_ptr != self.context.source.as_ptr() as usize
            || handle.source_len != self.context.source.len()
            || handle.index < self.index || handle.index >= self.node.subtree_end
            || handle.index >= self.context.nodes.len() { return None; }
        let node = self.context.nodes.get(handle.index);
        if node.rule != handle.rule || node.span != handle.span
            || R::from_capture(CaptureRule::decode(node.rule)).is_none() { return None; }
        Some(Self { context: self.context.clone(), index: handle.index, node, rule: PhantomData })
    }
    pub fn line_col(&self) -> (usize, usize) {
        self.as_span().start_pos().line_col()
    }
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
    /// Move this handle to its child when it has exactly one immediate child.
    /// A leaf or sibling boundary leaves it unchanged. Reuses the shared
    /// context without creating pair/iterator handles or incrementing its Rc.
    pub fn descend_single_child(&mut self) -> bool {
        let index = self.index + 1;
        if index >= self.node.subtree_end { return false; }
        let node = self.context.nodes.get(index);
        if node.subtree_end != self.node.subtree_end { return false; }
        self.index = index;
        self.node = node;
        true
    }
    /// Descend through selected wrappers with exactly one immediate child.
    /// Reuses this handle's context without cloning handles along the chain.
    /// Stops at an unselected rule, an empty node, or a sibling boundary.
    pub fn into_single_descendant_while<F>(mut self, mut transparent: F) -> Self
    where
        F: FnMut(R) -> bool,
    {
        while transparent(self.as_rule()) && self.descend_single_child() {}
        self
    }
    /// Select descendants in source order, skipping the subtree of each match.
    /// This excludes the current pair. Unselected captures are scanned in the
    /// contiguous arena without constructing handles or a traversal stack.
    /// Use this for declarations whose contents need a separate scoped visit.
    pub fn outermost_descendants<F>(self, mut select: F) -> impl Iterator<Item = Pair<'i, R>>
    where
        F: FnMut(R) -> bool,
    {
        let mut next = self.index + 1;
        let end = self.node.subtree_end;
        let context = self.context;
        std::iter::from_fn(move || {
            while next < end {
                let index = next;
                let node = context.nodes.get(index);
                next += 1;
                let rule = R::from_capture(CaptureRule::decode(node.rule))
                    .expect("capture rule missing from generated enum");
                if select(rule) {
                    next = node.subtree_end;
                    return Some(Pair {
                        context: context.clone(), index, node, rule: PhantomData,
                    });
                }
            }
            None
        })
    }
}
impl<'i, R: RuleIdentity> Pairs<'i, R> {
    /// Visit every remaining capture in preorder, including descendants.
    /// The contiguous arena requires no recursion or traversal stack.
    pub fn flatten(self) -> impl Iterator<Item = Pair<'i, R>> {
        let Self { context, mut next, end, .. } = self;
        std::iter::from_fn(move || {
            if next >= end { return None; }
            let index = next;
            next += 1;
            Some(Pair { node: context.nodes.get(index), context: context.clone(), index, rule: PhantomData })
        })
    }
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
    pub fn get_input(self) -> &'i str {
        self.source
    }
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
            let selected: Vec<_> = root.clone().outermost_descendants(|_| true).collect();
            assert_eq!(selected.len(), 1);
            assert_eq!(selected[0].as_rule(), Identity(CaptureRule::EndOfInput));
            assert_eq!(cloned.as_rule(), root.as_rule());
            let mut children = cloned.into_inner();
            assert_eq!(pairs.next().unwrap().as_rule(), root.as_rule());
            assert!(pairs.peek().is_none());
            drop(pairs);
            drop(root);
            let eof = children.peek().unwrap();
            assert_eq!(eof.as_rule(), Identity(CaptureRule::EndOfInput));
            assert_eq!(eof.as_str(), "");
            assert_eq!(eof.as_span().get_input(), "a");
            assert_eq!(eof.line_col(), (1, 2));
            assert_eq!((eof.as_span().start(), eof.as_span().end()), (1, 1));
            assert_eq!(children.next().unwrap().as_rule(), eof.as_rule());
            assert!(children.next().is_none());
        }
    }

    #[test]
    fn single_descendant_cursor_preserves_boundaries_spans_and_context() {
        let grammar = crate::compile(r#"
root = { wrapper ~ EOI }
wrapper = { inner }
inner = { leaf }
leaf = { "a" }
empty = { "" }
"#).unwrap();
        for wide in [false, true] {
            let mut tree = grammar.parse_source("root", "a").unwrap();
            if wide {
                let mut nodes = Nodes::Wide(Vec::new());
                for index in 0..tree.nodes.len() { nodes.push(tree.nodes.get(index)); }
                tree.nodes = nodes;
            }
            let mut root = tree.into_typed_pairs::<Identity>().next().unwrap();
            let context = Rc::as_ptr(&root.context);
            let references = Rc::strong_count(&root.context);
            assert!(!root.descend_single_child());
            assert_eq!(Rc::strong_count(&root.context), references);
            let unchanged = root.clone().into_single_descendant_while(|_| true);
            assert_eq!(unchanged.as_rule(), root.as_rule()); // EOI is a sibling.
            let wrapper = root.into_inner().next().unwrap();
            let inner_rule = Identity(CaptureRule::User(grammar.rule_id("inner").unwrap()));
            let mut cursor = wrapper.clone();
            let references = Rc::strong_count(&cursor.context);
            assert!(cursor.descend_single_child());
            assert_eq!(cursor.as_rule(), inner_rule);
            assert_eq!(Rc::strong_count(&cursor.context), references);
            assert!(cursor.descend_single_child());
            assert!(!cursor.descend_single_child());
            assert_eq!(Rc::as_ptr(&cursor.context), context);
            drop(cursor);
            let inner = wrapper.clone().into_single_descendant_while(|r| r != inner_rule);
            assert_eq!(inner.as_rule(), inner_rule);
            let leaf = wrapper.into_single_descendant_while(|_| true);
            assert_eq!(leaf.as_rule(), Identity(CaptureRule::User(grammar.rule_id("leaf").unwrap())));
            assert_eq!(leaf.as_str(), "a");
            assert_eq!((leaf.as_span().start(), leaf.as_span().end()), (0, 1));
            assert_eq!(Rc::as_ptr(&leaf.context), context);
            assert!(leaf.into_inner().next().is_none());
        }
        let empty = grammar.parse_source("empty", "").unwrap()
            .into_typed_pairs::<Identity>().next().unwrap();
        assert_eq!(empty.clone().into_single_descendant_while(|_| true).as_rule(), empty.as_rule());
    }

    #[test]
    fn filtered_descendants_preserve_order_prune_matches_and_exclude_root() {
        let grammar = crate::compile(r#"
root = { group ~ leaf ~ EOI }
group = { wrapper ~ leaf }
wrapper = { leaf }
leaf = { "a" }
"#).unwrap();
        let root = grammar.parse_source("root", "aaa").unwrap()
            .into_typed_pairs::<Identity>().next().unwrap();
        let flattened: Vec<_> = root.clone().into_inner().flatten()
            .map(|pair| (pair.as_rule(), pair.as_span().start())).collect();
        let named = |name| Identity(CaptureRule::User(grammar.rule_id(name).unwrap()));
        assert_eq!(flattened, [(named("group"), 0), (named("wrapper"), 0),
            (named("leaf"), 0), (named("leaf"), 1), (named("leaf"), 2),
            (Identity(CaptureRule::EndOfInput), 3)]);
        let group = Identity(CaptureRule::User(grammar.rule_id("group").unwrap()));
        let leaf = Identity(CaptureRule::User(grammar.rule_id("leaf").unwrap()));
        let selected: Vec<_> = root.clone().outermost_descendants(|r| r == group || r == leaf)
            .map(|p| (p.as_rule(), p.as_str(), p.as_span().start())).collect();
        assert_eq!(selected, [(group, "aa", 0), (leaf, "a", 2)]);
        let leaves: Vec<_> = root.clone().outermost_descendants(|r| r == leaf)
            .map(|p| p.as_span().start()).collect();
        assert_eq!(leaves, [0, 1, 2]);
        assert!(root.clone().outermost_descendants(|r| r == root.as_rule()).next().is_none());
        assert!(root.outermost_descendants(|_| false).next().is_none());
    }
}
