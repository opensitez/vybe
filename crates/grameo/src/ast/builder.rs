//! Owned AST assembly driven by generated binding tables.
use super::{Binding, Cardinality, Field, Source};
use crate::{
    Span,
    arena::{Id, MetadataArena},
    builder::{BuildError, Builder},
};

pub enum Argument<'a, V> {
    One(V),
    Optional(Option<V>),
    Many(Vec<V>),
    Text(&'a str),
    Span(Span),
}

/// Constructors must have no external effects and must not panic. `split`
/// restores the exact owned children, flattened in field order. `Undo` can
/// retain compact inverse data for lowering that changes an operand.
pub trait Target {
    type Value;
    type Undo;
    type Context;
    /// Prepared once per input, e.g. a source line index. It is immutable during
    /// speculation; program scopes and policies belong in constructed values.
    fn prepare(source: &str) -> Self::Context;
    fn validate(
        _: &Self::Context,
        _: usize,
        _: &[Field],
        _: Span,
        _: &str,
    ) -> Result<(), BuildError> {
        Ok(())
    }
    fn construct(
        context: &Self::Context,
        constructor: usize,
        span: Span,
        arguments: Vec<Argument<'_, Self::Value>>,
    ) -> (Self::Value, Self::Undo);
    fn split(constructor: usize, value: Self::Value, undo: Self::Undo) -> Vec<Self::Value>;
}

#[derive(Clone, Copy)]
struct Meta {
    binding: &'static Binding,
}
enum Node<V> {
    Frame(usize),
    Ready(Meta, V),
}
struct Undo<U> {
    children: Vec<Meta>,
    target: U,
}

pub struct AstBuilder<T: Target> {
    bindings: Vec<Option<&'static Binding>>,
    markers: Vec<Option<Id>>,
    arena: MetadataArena<Node<T::Value>, Undo<T::Undo>>,
    context: T::Context,
}
impl<T: Target> AstBuilder<T> {
    pub fn new(bindings: &'static [Binding], source: &str) -> Self {
        Self::with_context(bindings, T::prepare(source))
    }
    /// Supply immutable per-input context before matching, including frontend
    /// policy required by reductions. The context must describe this input;
    /// speculative constructors must not mutate external scope state.
    pub fn with_context(bindings: &'static [Binding], context: T::Context) -> Self {
        let mut by_rule = vec![None; bindings.iter().map(|b| b.rule).max().map_or(0, |n| n + 1)];
        for binding in bindings {
            by_rule[binding.rule] = Some(binding);
        }
        let mut arena = MetadataArena::default();
        let markers = by_rule
            .iter()
            .enumerate()
            .map(|(rule, binding)| binding.map(|_| arena.alloc(Node::Frame(rule))))
            .collect();
        // Immutable frame sentinels are shared across nested/repeated matches.
        // Their allocations precede every parser checkpoint.
        arena.commit();
        Self {
            bindings: by_rule,
            markers,
            arena,
            context,
        }
    }
    pub fn into_value(self) -> Result<T::Value, BuildError> {
        if self.arena.values().len() != 1 {
            return Err(BuildError {
                span: Span { start: 0, end: 0 },
                message: "AST root must construct exactly one value".into(),
            });
        }
        let root = self.arena.values()[0];
        match self.arena.into_root(root) {
            Node::Ready(_, value) => Ok(value),
            Node::Frame(_) => unreachable!("unfinished AST frame after successful parse"),
        }
    }
    fn binding(&self, rule: usize) -> Option<&'static Binding> {
        self.bindings.get(rule).copied().flatten()
    }
}
fn split<T: Target>(node: Node<T::Value>, undo: Undo<T::Undo>) -> Vec<Node<T::Value>> {
    let Node::Ready(meta, value) = node else {
        unreachable!()
    };
    let values = T::split(meta.binding.constructor, value, undo.target);
    assert_eq!(
        values.len(),
        undo.children.len(),
        "AST constructor inverse changed child count"
    );
    undo.children
        .into_iter()
        .zip(values)
        .map(|(meta, value)| Node::Ready(meta, value))
        .collect()
}
impl<T: Target> Builder for AstBuilder<T> {
    fn observes_rule(&self, rule: usize) -> bool {
        self.binding(rule).is_some()
    }
    fn checkpoint(&self) -> usize {
        self.arena.checkpoint()
    }
    fn rollback(&mut self, mark: usize) {
        self.arena.rollback(mark);
    }
    fn committed(&mut self, _: &str, _: usize) {
        self.arena.commit();
    }
    fn begin(&mut self, rule: usize, _: usize) {
        self.arena
            .push(self.markers[rule].expect("observed AST frame"));
    }
    fn finish(&mut self, _: usize, _: Span, _: &str) {
        unreachable!("checked AST finish is required")
    }
    fn try_finish(&mut self, rule: usize, span: Span, source: &str) -> Result<(), BuildError> {
        let binding = self.binding(rule).expect("observed AST binding");
        let fail = |message: &str| BuildError {
            span,
            message: message.into(),
        };
        let mut children = Vec::new();
        loop {
            let id = self.arena.pop().ok_or_else(|| fail("missing AST frame"))?;
            match self.arena.get(id) {
                Node::Frame(frame) if *frame == rule => break,
                Node::Frame(_) => return Err(fail("unfinished child AST frame")),
                Node::Ready(_, _) => children.push(id),
            }
        }
        children.reverse();
        let mut selected = Vec::with_capacity(children.len());
        let mut counts = Vec::with_capacity(binding.fields.len());
        let mut used = vec![false; children.len()];
        for field in binding.fields {
            let (rule, occurrence, cardinality) = match field.source {
                Source::Child {
                    rule,
                    occurrence,
                    cardinality,
                } => (Some(rule), occurrence, cardinality),
                Source::Children(cardinality) => (None, None, cardinality),
                _ => {
                    counts.push(0);
                    continue;
                }
            };
            let mut count = 0;
            let mut ordinal = 0;
            for (index, &child) in children.iter().enumerate() {
                let Node::Ready(meta, _) = self.arena.get(child) else {
                    unreachable!()
                };
                if rule.is_some_and(|rule| meta.binding.rule != rule) {
                    continue;
                }
                let take = occurrence.is_none_or(|wanted| wanted == ordinal);
                ordinal += 1;
                if take {
                    if used[index] {
                        return Err(fail("AST child selected by multiple fields"));
                    }
                    used[index] = true;
                    selected.push(child);
                    count += 1;
                }
            }
            if match cardinality {
                Cardinality::One => count != 1,
                Cardinality::Optional => count > 1,
                Cardinality::Many => false,
            } {
                return Err(fail(
                    "AST child cardinality does not match its declared field",
                ));
            }
            counts.push(count);
        }
        if used.iter().any(|used| !used) {
            return Err(fail("AST binding leaves a semantic child unconsumed"));
        }
        let text = &source[span.start..span.end];
        T::validate(
            &self.context,
            binding.constructor,
            binding.fields,
            span,
            text,
        )?;
        let context = &self.context;
        let node = self.arena.list_with_metadata(
            selected,
            |nodes| {
                let mut metadata = Vec::with_capacity(nodes.len());
                let mut values = nodes.into_iter().map(|node| {
                    let Node::Ready(meta, value) = node else {
                        unreachable!()
                    };
                    metadata.push(meta);
                    value
                });
                let mut arguments = Vec::with_capacity(binding.fields.len());
                for (field, count) in binding.fields.iter().zip(counts) {
                    arguments.push(match field.source {
                        Source::Text => Argument::Text(text),
                        Source::Span => Argument::Span(span),
                        Source::Children(Cardinality::One)
                        | Source::Child {
                            cardinality: Cardinality::One,
                            ..
                        } => Argument::One(values.next().unwrap()),
                        Source::Children(Cardinality::Optional)
                        | Source::Child {
                            cardinality: Cardinality::Optional,
                            ..
                        } => Argument::Optional((count == 1).then(|| values.next().unwrap())),
                        Source::Children(Cardinality::Many)
                        | Source::Child {
                            cardinality: Cardinality::Many,
                            ..
                        } => Argument::Many(values.by_ref().take(count).collect()),
                    });
                }
                assert!(values.next().is_none());
                let (value, target) = T::construct(context, binding.constructor, span, arguments);
                (
                    Node::Ready(Meta { binding }, value),
                    Undo {
                        children: metadata,
                        target,
                    },
                )
            },
            split::<T>,
        );
        self.arena.push(node);
        Ok(())
    }
    fn prune_capture_children(&self, rule: usize, _: Span) -> bool {
        self.binding(rule).is_some()
    }
}
