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
/// Does not yet certify progress/recursion safety or supply a source parser.
#[derive(Debug, Clone)]
pub struct CompiledGrammar {
    syntax: GrammarSyntax,
    references: Vec<Option<Reference>>,
    rules_by_name: HashMap<String, RuleId>,
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
}

pub fn compile(source: &str) -> Result<CompiledGrammar, Vec<Diagnostic>> {
    resolve(crate::parse(source).map_err(|error| vec![error])?)
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
        Ok(CompiledGrammar {
            syntax,
            references,
            rules_by_name,
        })
    } else {
        Err(diagnostics)
    }
}
