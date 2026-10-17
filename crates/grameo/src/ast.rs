//! Declarative, target-independent AST construction bindings.
//! Annotations are ordinary PEG comments; constructors belong to the target.
use crate::{CompiledGrammar, Diagnostic, Span};
use std::collections::HashSet;

pub mod builder;

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Cardinality {
    One,
    Optional,
    Many,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub enum SourceSyntax {
    Children(Cardinality),
    Child {
        rule: String,
        occurrence: Option<usize>,
        cardinality: Cardinality,
    },
    Text,
    Span,
}
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct FieldSyntax {
    pub name: String,
    pub source: SourceSyntax,
}
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct BindingSyntax {
    pub rule: String,
    pub constructor: String,
    pub fields: Vec<FieldSyntax>,
    pub span: Span,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Source {
    Children(Cardinality),
    Child {
        rule: usize,
        occurrence: Option<usize>,
        cardinality: Cardinality,
    },
    Text,
    Span,
}
#[derive(Debug, Clone, Copy)]
pub struct Field {
    pub name: &'static str,
    pub source: Source,
}
/// Static data emitted by codegen. Constructor IDs index its generated enum.
#[derive(Debug)]
pub struct Binding {
    pub rule: usize,
    pub constructor: usize,
    pub fields: &'static [Field],
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum FieldKind<'a> {
    Child {
        value_kind: &'a str,
        cardinality: Cardinality,
    },
    Text,
    Span,
}
pub struct FieldSchema<'a> {
    pub name: &'a str,
    pub kind: FieldKind<'a>,
}
pub struct ConstructorSchema<'a> {
    pub name: &'a str,
    pub value_kind: &'a str,
    pub fields: &'a [FieldSchema<'a>],
}
impl CompiledGrammar {
    /// Check constructor names, field signatures and child value types against
    /// the embedding's schema at build time, without depending on its AST crate.
    pub fn validate_ast_schema(
        &self,
        schema: &[ConstructorSchema<'_>],
    ) -> Result<(), Vec<Diagnostic>> {
        let mut errors = Vec::new();
        let mut names = HashSet::new();
        for constructor in schema {
            if !names.insert(constructor.name) {
                errors.push(Diagnostic::new(
                    "G042",
                    format!("duplicate target AST constructor '{}'", constructor.name),
                    Span { start: 0, end: 0 },
                ));
            }
        }
        for binding in &self.syntax().ast_bindings {
            let error = |message: String| Diagnostic::new("G042", message, binding.span);
            let Some(constructor) = schema.iter().find(|c| c.name == binding.constructor) else {
                errors.push(error(format!(
                    "unknown target AST constructor '{}'",
                    binding.constructor
                )));
                continue;
            };
            if constructor.fields.len() != binding.fields.len() {
                errors.push(error(
                    "AST constructor field count differs from target schema".into(),
                ));
                continue;
            }
            for (field, expected) in binding.fields.iter().zip(constructor.fields) {
                let actual = match &field.source {
                    SourceSyntax::Children(cardinality) => {
                        let mut kinds = child_bindings(self, self.rule_id(&binding.rule).unwrap())
                            .into_iter()
                            .filter_map(|rule| {
                                self.syntax()
                                    .ast_bindings
                                    .iter()
                                    .find(|b| self.rule_id(&b.rule) == Some(rule))
                            })
                            .filter_map(|b| {
                                schema
                                    .iter()
                                    .find(|c| c.name == b.constructor)
                                    .map(|c| c.value_kind)
                            });
                        let kind = kinds.next();
                        if kinds.all(|other| Some(other) == kind) {
                            kind.map(|value_kind| FieldKind::Child {
                                value_kind,
                                cardinality: *cardinality,
                            })
                        } else {
                            None
                        }
                    }
                    SourceSyntax::Text => Some(FieldKind::Text),
                    SourceSyntax::Span => Some(FieldKind::Span),
                    SourceSyntax::Child {
                        rule, cardinality, ..
                    } => self
                        .syntax()
                        .ast_bindings
                        .iter()
                        .find(|b| &b.rule == rule)
                        .and_then(|b| schema.iter().find(|c| c.name == b.constructor))
                        .map(|c| FieldKind::Child {
                            value_kind: c.value_kind,
                            cardinality: *cardinality,
                        }),
                };
                if field.name != expected.name || actual != Some(expected.kind) {
                    errors.push(error(format!(
                        "AST field '{}' does not match target constructor '{}'",
                        field.name, binding.constructor
                    )));
                }
            }
        }
        if errors.is_empty() {
            Ok(())
        } else {
            Err(errors)
        }
    }
}

pub(crate) fn parse_binding(text: &str, span: Span) -> Result<BindingSyntax, Diagnostic> {
    struct Tokens<'a> {
        text: &'a str,
        offset: usize,
        span: Span,
    }
    impl Tokens<'_> {
        fn fail(&self, message: &str) -> Diagnostic {
            Diagnostic::new("G040", message, self.span)
        }
        fn space(&mut self) {
            while self
                .text
                .as_bytes()
                .get(self.offset)
                .is_some_and(u8::is_ascii_whitespace)
            {
                self.offset += 1;
            }
        }
        fn eat(&mut self, ch: u8) -> bool {
            self.space();
            if self.text.as_bytes().get(self.offset) == Some(&ch) {
                self.offset += 1;
                true
            } else {
                false
            }
        }
        fn expect(&mut self, ch: u8) -> Result<(), Diagnostic> {
            if self.eat(ch) {
                Ok(())
            } else {
                Err(self.fail("invalid AST binding punctuation"))
            }
        }
        fn ident(&mut self) -> Result<String, Diagnostic> {
            self.space();
            let start = self.offset;
            if !self
                .text
                .as_bytes()
                .get(start)
                .is_some_and(|c| c.is_ascii_alphabetic() || *c == b'_')
            {
                return Err(self.fail("expected AST binding identifier"));
            }
            self.offset += 1;
            while self
                .text
                .as_bytes()
                .get(self.offset)
                .is_some_and(|c| c.is_ascii_alphanumeric() || *c == b'_')
            {
                self.offset += 1;
            }
            Ok(self.text[start..self.offset].to_owned())
        }
    }
    let mut t = Tokens {
        text,
        offset: 0,
        span,
    };
    let rule = t.ident()?;
    t.expect(b'=')?;
    let constructor = t.ident()?;
    t.expect(b'(')?;
    let mut fields = Vec::new();
    if !t.eat(b')') {
        loop {
            let name = t.ident()?;
            t.expect(b':')?;
            let selector = t.ident()?;
            let source = match selector.as_str() {
                "source" => SourceSyntax::Text,
                "rule" => SourceSyntax::Span,
                "children" => SourceSyntax::Children(if t.eat(b'?') {
                    Cardinality::Optional
                } else if t.eat(b'*') {
                    Cardinality::Many
                } else {
                    Cardinality::One
                }),
                _ => {
                    let occurrence = if t.eat(b'[') {
                        t.space();
                        let start = t.offset;
                        while t
                            .text
                            .as_bytes()
                            .get(t.offset)
                            .is_some_and(u8::is_ascii_digit)
                        {
                            t.offset += 1;
                        }
                        let index = t.text[start..t.offset]
                            .parse()
                            .map_err(|_| t.fail("invalid child occurrence index"))?;
                        t.expect(b']')?;
                        Some(index)
                    } else {
                        None
                    };
                    let cardinality = if t.eat(b'?') {
                        Cardinality::Optional
                    } else if t.eat(b'*') {
                        Cardinality::Many
                    } else {
                        Cardinality::One
                    };
                    if occurrence.is_some() && cardinality == Cardinality::Many {
                        return Err(t.fail("indexed children cannot use '*'"));
                    }
                    SourceSyntax::Child {
                        rule: selector,
                        occurrence,
                        cardinality,
                    }
                }
            };
            fields.push(FieldSyntax { name, source });
            if t.eat(b')') {
                break;
            }
            t.expect(b',')?;
        }
    }
    t.space();
    if t.offset != text.len() {
        return Err(t.fail("trailing AST binding text"));
    }
    Ok(BindingSyntax {
        rule,
        constructor,
        fields,
        span,
    })
}

/// Validate names and ownership selectors before any parser is generated.
fn child_bindings(grammar: &CompiledGrammar, rule: usize) -> Vec<usize> {
    let mut pending = vec![grammar.syntax().rules[rule].expression];
    let mut seen = HashSet::new();
    let mut children = HashSet::new();
    while let Some(id) = pending.pop() {
        if !seen.insert(id) {
            continue;
        }
        match grammar.syntax().expressions[id].kind {
            crate::ExprKind::Reference(_) => {
                if let Some(crate::Reference::Rule(child)) = grammar.reference(id) {
                    if grammar
                        .syntax()
                        .ast_bindings
                        .iter()
                        .any(|b| grammar.rule_id(&b.rule) == Some(child))
                    {
                        children.insert(child);
                    } else {
                        pending.push(grammar.syntax().rules[child].expression);
                    }
                }
            }
            crate::ExprKind::Sequence { left, right } | crate::ExprKind::Choice { left, right } => {
                pending.push(left);
                pending.push(right);
            }
            crate::ExprKind::Repeat { expression, .. }
            | crate::ExprKind::Group(expression)
            | crate::ExprKind::Push(expression) => pending.push(expression),
            _ => {} // Predicates and implicit trivia never construct values.
        }
    }
    let mut children = children.into_iter().collect::<Vec<_>>();
    children.sort_unstable();
    children
}

pub(crate) fn validate(grammar: &CompiledGrammar) -> Vec<Diagnostic> {
    let mut errors = Vec::new();
    let mut bound = HashSet::new();
    for binding in &grammar.syntax().ast_bindings {
        let error = |message: String| Diagnostic::new("G041", message, binding.span);
        if grammar.rule_id(&binding.rule).is_none() {
            errors.push(error(format!(
                "AST binding refers to unknown rule '{}'",
                binding.rule
            )));
        }
        if !bound.insert(&binding.rule) {
            errors.push(error(format!(
                "duplicate AST binding for '{}'",
                binding.rule
            )));
        }
        let mut names = HashSet::new();
        let mut selectors = HashSet::new();
        for field in &binding.fields {
            if !names.insert(&field.name) {
                errors.push(error(format!("duplicate AST field '{}'", field.name)));
            }
            if let SourceSyntax::Child {
                rule, occurrence, ..
            } = &field.source
            {
                if grammar.rule_id(rule).is_none() {
                    errors.push(error(format!("unknown AST child rule '{rule}'")));
                }
                if !selectors.insert((rule, occurrence)) {
                    errors.push(error(format!(
                        "AST child '{rule}' is consumed more than once"
                    )));
                }
            }
        }
    }
    for binding in &grammar.syntax().ast_bindings {
        let reachable = grammar
            .rule_id(&binding.rule)
            .map(|rule| child_bindings(grammar, rule))
            .unwrap_or_default();
        for field in &binding.fields {
            if let SourceSyntax::Child { rule, .. } = &field.source
                && !bound.contains(rule)
            {
                errors.push(Diagnostic::new(
                    "G041",
                    format!("AST child '{rule}' has no construction binding"),
                    binding.span,
                ));
            }
            if let SourceSyntax::Child { rule, .. } = &field.source
                && grammar
                    .rule_id(rule)
                    .is_some_and(|rule| !reachable.contains(&rule))
            {
                errors.push(Diagnostic::new(
                    "G041",
                    format!(
                        "AST child '{rule}' is not a semantic child of '{}'",
                        binding.rule
                    ),
                    binding.span,
                ));
            }
        }
    }
    errors
}
