use crate::grammar::{ExprId, RuleId};
use crate::{Diagnostic, ExprKind, GrammarSyntax};
use std::collections::HashMap;

/// Supported builtin inventory covers the repository grammars, including source
/// recognition. Other Unicode properties remain explicit unsupported features.
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
    pub(crate) empty_success_heads:HashMap<RuleId,crate::program::EmptySuccessHead>,
    pub(crate) empty_suffix_heads:HashMap<usize,crate::program::EmptySuffixHead>,
    pub(crate) shared_inner_prefixes: HashMap<usize,crate::program::SharedInnerPrefix>,
    pub(crate) shared_group_prefixes: HashMap<usize,crate::program::SharedGroupPrefix>,
    pub(crate) shared_wrapper_prefixes: HashMap<usize, crate::program::SharedWrapperPrefix>,
    pub(crate) shared_prefixes: HashMap<usize, crate::program::SharedPrefixChoice>,
    pub(crate) pratt: Vec<Option<crate::program::OwnedPratt>>,
    pub(crate) trivia_failure: Vec<Option<crate::lexical::FailurePrefix>>,
    pub(crate) leading_bytes: Vec<Option<crate::lexical::LeadingBytes>>,
    pub(crate) atomic_expressions: Vec<bool>,
    pub(crate) terminal_choices: Vec<Option<crate::lexical::OwnedTerminalChoice>>,
    pub(crate) failure_terminals: Vec<Vec<usize>>,
    pub(crate) literal_choices: HashMap<usize, crate::lexical::LiteralChoice>,
    literal_choice_failures: bool,
    pub(crate) choice_head_rules: Vec<RuleId>,
    pub(crate) owned_capture_boundaries: HashMap<RuleId,Vec<RuleId>>,
    pub(crate) source_units: Vec<crate::source_units::OwnedBinding>,
    pub(crate) code_space: Option<crate::source_units::CodeSpace>,
    pub(crate) optimization_selection: crate::optimization::OptimizationSelection,
    pub(crate) native_token_compilation: bool,
    pub(crate) native_token_failures: bool,
}

impl CompiledGrammar {
    /// Experimental pure literal-choice failure dispatch. Disabled by default:
    /// current full AST benchmarks do not establish a consistent performance
    /// win. Disabling also releases the certificate table.
    pub fn set_literal_choice_failures(&mut self, enabled: bool) {
        self.literal_choice_failures = enabled;
        self.literal_choices = if enabled {
            crate::lexical::literal_choices(self)
        } else {
            HashMap::new()
        };
    }
    /// Maximal pure literal-choice certificates in deterministic expression order.
    pub fn literal_choices(&self) -> impl Iterator<Item = (usize, &crate::lexical::LiteralChoice)> {
        let mut choices: Vec<_> = self
            .literal_choices
            .iter()
            .map(|(&id, choice)| (id, choice))
            .collect();
        choices.sort_unstable_by_key(|(id, _)| *id);
        choices.into_iter()
    }
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
        self.failure_terminals = crate::lexical::failure_terminals(self);
        self.terminal_choices = crate::lexical::terminal_choices(self);
        self.atomic_expressions = crate::lexical::atomic_expressions(self);
        self.leading_bytes = crate::lexical::leading_bytes(self);
        if self.literal_choice_failures {
            self.literal_choices = crate::lexical::literal_choices(self);
        }
    }
}

/// Explicit grammar annotation activation. Ordinary compile leaves both disabled.
#[derive(Debug, Clone, Copy, Default)]
pub struct CompileOptions {
    /// Read and validate @grameo ast constructor bindings.
    pub ast_bindings: bool,
    /// Read @grameo optimize, pratt and operator performance directives.
    pub optimizations: bool,
}

pub fn compile_with_options(source: &str, options: CompileOptions) -> Result<CompiledGrammar, Vec<Diagnostic>> {
    let (syntax, annotations) = crate::grammar::parse_configured(source, crate::ParseOptions::default(), options.ast_bindings, options.optimizations).map_err(|e| vec![e])?;
    let mut grammar = resolve(syntax)?;
    if !grammar.analysis.errors.is_empty() {
        return Err(grammar.analysis.errors.clone());
    }
    crate::optimization::apply(&mut grammar, annotations)?;
    Ok(grammar)
}

pub fn compile(source: &str) -> Result<CompiledGrammar, Vec<Diagnostic>> {
    let grammar = resolve(crate::parse(source).map_err(|error| vec![error])?)?;
    if grammar.analysis.errors.is_empty() {
        Ok(grammar)
    } else {
        Err(grammar.analysis.errors.clone())
    }
}

/// Compile PEG with explicitly enabled AST annotation validation.
pub fn compile_with_ast_bindings(source: &str) -> Result<CompiledGrammar, Vec<Diagnostic>> {
    let grammar = resolve(crate::grammar::parse_with_ast_bindings(source).map_err(|error| vec![error])?)?;
    if grammar.analysis.errors.is_empty() { Ok(grammar) }
    else { Err(grammar.analysis.errors.clone()) }
}

pub fn resolve(mut syntax: GrammarSyntax) -> Result<CompiledGrammar, Vec<Diagnostic>> {
    crate::repetition_tails::normalize(&mut syntax);
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
            shared_prefixes: HashMap::new(),
            empty_success_heads: HashMap::new(),
            empty_suffix_heads: HashMap::new(),
            shared_inner_prefixes: HashMap::new(),
            shared_group_prefixes: HashMap::new(),
            shared_wrapper_prefixes: HashMap::new(),
            pratt: vec![None; syntax.rules.len()],
            trivia_failure: vec![None; syntax.rules.len()],
            failure_terminals: vec![Vec::new(); syntax.rules.len()],
            terminal_choices: Vec::new(),
            atomic_expressions: Vec::new(),
            leading_bytes: Vec::new(),
            literal_choices: HashMap::new(),
            literal_choice_failures: false,
            choice_head_rules: Vec::new(),
            owned_capture_boundaries: HashMap::new(),
            source_units: Vec::new(),
            code_space: None,
            optimization_selection: crate::optimization::OptimizationSelection::default(),
            native_token_compilation: false,
            native_token_failures: false,
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
        let ast_errors = crate::ast::validate(&grammar);
        if !ast_errors.is_empty() { return Err(ast_errors); }
        Ok(grammar)
    } else {
        Err(diagnostics)
    }
}
