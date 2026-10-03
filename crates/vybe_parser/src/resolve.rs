use crate::grammar::{ExprId, RuleId};
use crate::{Diagnostic, ExprKind, GrammarSyntax};
use std::collections::HashMap;

/// Initial builtin inventory covers the repository grammars, not every Unicode
/// property pest offers. Recognition of these builtins is a later milestone.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Builtin {
    Any,
    Soi,
    Eoi,
    Ascii,
    AsciiDigit,
    AsciiNonzeroDigit,
    AsciiBinDigit,
    AsciiOctDigit,
    AsciiHexDigit,
    AsciiAlphaLower,
    AsciiAlphaUpper,
    AsciiAlpha,
    AsciiAlphanumeric,
    Newline,
    XidStart,
    XidContinue,
    Peek,
    PeekAll,
    Pop,
    PopAll,
    Drop,
}

impl Builtin {
    pub fn from_name(name: &str) -> Option<Self> {
        Some(match name {
            "ANY" => Self::Any,
            "SOI" => Self::Soi,
            "EOI" => Self::Eoi,
            "ASCII" => Self::Ascii,
            "ASCII_DIGIT" => Self::AsciiDigit,
            "ASCII_NONZERO_DIGIT" => Self::AsciiNonzeroDigit,
            "ASCII_BIN_DIGIT" => Self::AsciiBinDigit,
            "ASCII_OCT_DIGIT" => Self::AsciiOctDigit,
            "ASCII_HEX_DIGIT" => Self::AsciiHexDigit,
            "ASCII_ALPHA_LOWER" => Self::AsciiAlphaLower,
            "ASCII_ALPHA_UPPER" => Self::AsciiAlphaUpper,
            "ASCII_ALPHA" => Self::AsciiAlpha,
            "ASCII_ALPHANUMERIC" => Self::AsciiAlphanumeric,
            "NEWLINE" => Self::Newline,
            "XID_START" => Self::XidStart,
            "XID_CONTINUE" => Self::XidContinue,
            "PEEK" => Self::Peek,
            "PEEK_ALL" => Self::PeekAll,
            "POP" => Self::Pop,
            "POP_ALL" => Self::PopAll,
            "DROP" => Self::Drop,
            _ => return None,
        })
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Reference {
    Rule(RuleId),
    Builtin(Builtin),
}

/// Located grammar IR with names resolved once to dense numeric IDs.
/// Includes conservative progress/effect analysis. Warnings require runtime
/// guards; they do not prove a grammar unsafe or cause compilation rejection.
#[derive(Debug, Clone)]
pub struct CompiledGrammar {
    syntax: GrammarSyntax,
    references: Vec<Option<Reference>>,
    rules_by_name: HashMap<String, RuleId>,
    analysis: crate::analysis::Analysis,
    pub(crate) scopes: Vec<crate::program::Scope>,
    pub(crate) rule_scopes: Vec<usize>,
    pub(crate) pratt: Vec<Option<crate::program::OwnedPratt>>,
    pub(crate) trivia_failure: Vec<Option<crate::lexical::FailurePrefix>>,
}

impl CompiledGrammar {
    pub fn syntax(&self) -> &GrammarSyntax {
        &self.syntax
    }
    pub fn rule_id(&self, name: &str) -> Option<RuleId> {
        self.rules_by_name.get(name).copied()
    }
    pub fn reference(&self, expression: ExprId) -> Option<Reference> {
        self.references.get(expression).copied().flatten()
    }
    pub fn analysis(&self) -> &crate::analysis::Analysis {
        &self.analysis
    }
    pub fn scopes(&self) -> &[crate::program::Scope] {
        &self.scopes
    }
    pub fn entries(&self) -> impl Iterator<Item = (&str, RuleId)> {
        self.rules_by_name
            .iter()
            .map(|(name, id)| (name.as_str(), *id))
    }
    pub(crate) fn add_entry(&mut self, name: String, id: RuleId) {
        self.rules_by_name.insert(name, id);
    }
    pub(crate) fn refresh_analysis(&mut self) {
        self.analysis = crate::analysis::analyze(self);
        self.trivia_failure = crate::lexical::trivia_failure_prefixes(self);
    }
}

pub fn compile(source: &str) -> Result<CompiledGrammar, Vec<Diagnostic>> {
    let grammar = resolve(crate::parse(source).map_err(|error| vec![error])?)?;
    if grammar.analysis.errors.is_empty() {
        Ok(grammar)
    } else {
        Err(grammar.analysis.errors.clone())
    }
}

pub fn resolve(syntax: GrammarSyntax) -> Result<CompiledGrammar, Vec<Diagnostic>> {
    let mut rules_by_name: HashMap<String, RuleId> = HashMap::with_capacity(syntax.rules.len());
    let mut diagnostics = Vec::new();
    for (id, rule) in syntax.rules.iter().enumerate() {
        if let Some(&previous) = rules_by_name.get(&rule.name) {
            let mut error = Diagnostic::new(
                "G009",
                format!("duplicate rule '{}'", rule.name),
                rule.name_span,
            );
            error
                .related
                .push((syntax.rules[previous].name_span, "first definition".into()));
            diagnostics.push(error);
        } else {
            rules_by_name.insert(rule.name.clone(), id);
        }
    }
    let mut references = vec![None; syntax.expressions.len()];
    for (id, expression) in syntax.expressions.iter().enumerate() {
        if let ExprKind::Reference(name) = &expression.kind {
            references[id] = rules_by_name
                .get(name)
                .copied()
                .map(Reference::Rule)
                .or_else(|| Builtin::from_name(name).map(Reference::Builtin));
            if references[id].is_none() {
                diagnostics.push(Diagnostic::new(
                    "G010",
                    format!("undefined rule or unsupported builtin '{name}'"),
                    expression.span,
                ));
            }
        }
    }
    if diagnostics.is_empty() {
        let mut grammar = CompiledGrammar {
            pratt: vec![None; syntax.rules.len()],
            trivia_failure: vec![None; syntax.rules.len()],
            rule_scopes: vec![0; syntax.rules.len()],
            syntax,
            references,
            rules_by_name,
            analysis: crate::analysis::Analysis::default(),
            scopes: vec![crate::program::Scope::default()],
        };
        let whitespace = grammar.rule_id("WHITESPACE");
        let (whitespace_class, whitespace_prefix_class) =
            crate::lexical::whitespace_classes(&grammar, whitespace);
        grammar.scopes[0] = crate::program::Scope {
            whitespace,
            comment: grammar.rule_id("COMMENT"),
            whitespace_class,
            whitespace_prefix_class,
        };
        grammar.refresh_analysis();
        Ok(grammar)
    } else {
        Err(diagnostics)
    }
}
