//! Opt-in completed-statement lowering while Grameo is advancing.
//!
//! The caller supplies an authoritative normalized manifest from semantic/header
//! discovery. Fresh contextual lowering runs at each completed top-level capture;
//! late/global fixups use that manifest as a per-statement fallback. Compiler
//! output remains unpublished until this producer and its parser both succeed.
use super::*;
use grameo::builder::{BuildError, Builder};

/// The owned compiler retains the normalized body itself; no replay facts escape.
struct OwnedSourceContext {
    original_source: std::sync::Arc<str>,
    committed_preparation: Option<crate::GrammarCommit>,
}
pub(super) fn owned_source_context(source: &str, commit: crate::GrammarCommit)
    -> Result<vybe_compiler::statement_codegen::ProducerContext, String> {
    validate_source_commit(source, Some(commit))?;
    Ok(std::sync::Arc::new(OwnedSourceContext {
        original_source: source.into(), committed_preparation: Some(commit),
    }))
}
fn validate_source_commit(source: &str, commit: Option<crate::GrammarCommit>) -> Result<(), String> {
    let commit = commit.ok_or("PHP prepared source has no completed normalization seal")?;
    if commit.source_bytes() != source.len() || commit.consumed() != commit.parser_bytes() {
        return Err("PHP prepared handoff has an inconsistent grammar commit".into());
    }
    Ok(())
}
#[cfg(feature = "service-verification")]
std::thread_local! { static FACTS_CAPTURES: std::cell::Cell<usize> = const { std::cell::Cell::new(0) }; }
#[cfg(feature = "service-verification")]
pub(crate) fn facts_capture_count() -> usize { FACTS_CAPTURES.with(|count| count.get()) }
#[cfg(feature = "service-verification")]
pub(crate) fn uncommitted_owned_context_rejected() -> bool {
    let context: vybe_compiler::statement_codegen::ProducerContext = std::sync::Arc::new(
        OwnedSourceContext { original_source: "".into(), committed_preparation: None });
    validate_prepared_context("", &context).is_err()
}

/// Immutable per-source discovery, never source-ordered final walker state.
pub(super) struct PhpDeclarationFacts {
    original_source: std::sync::Arc<str>,
    committed_preparation: Option<crate::GrammarCommit>,
    normalized_source: std::sync::Arc<str>,
    func_registry: std::collections::HashMap<String, FuncMeta>,
    func_by_ref_params: std::collections::HashMap<String, (String, Vec<bool>)>,
    namespace_functions: std::collections::HashSet<String>,
    type_kinds: std::collections::HashMap<String, &'static str>,
    predeclared_type_names: std::collections::HashSet<String>,
}
impl PhpDeclarationFacts {
    pub(super) fn capture(original: &str, normalized: String, walker: &PhpWalker) -> Self {
        #[cfg(feature = "service-verification")]
        FACTS_CAPTURES.with(|count| count.set(count.get() + 1));
        Self { original_source: original.into(), committed_preparation: None, normalized_source: normalized.into(),
            func_registry: walker.func_registry.clone(), func_by_ref_params: walker.func_by_ref_params.clone(),
            namespace_functions: walker.namespace_functions.clone(), type_kinds: walker.type_kinds.clone(),
            predeclared_type_names: walker.predeclared_type_names.clone() }
    }
    /// Called only after the original grammar commit AND all contextual
    /// normalization have succeeded, before this Arc escapes preparation.
    pub(super) fn seal(&mut self, commit: crate::GrammarCommit) -> Result<(), String> {
        if commit.source_bytes() != self.original_source.len()
            || commit.consumed() != commit.parser_bytes() {
            return Err("PHP prepared handoff has an inconsistent grammar commit".into());
        }
        if self.committed_preparation.is_some() {
            return Err("PHP prepared handoff was sealed twice".into());
        }
        self.committed_preparation = Some(commit);
        Ok(())
    }
    fn install(&self, walker: &mut PhpWalker) {
        walker.func_registry = self.func_registry.clone();
        walker.func_by_ref_params = self.func_by_ref_params.clone();
        walker.namespace_functions = self.namespace_functions.clone();
        walker.type_kinds = self.type_kinds.clone();
        walker.predeclared_type_names = self.predeclared_type_names.clone();
    }
}

#[derive(Clone, Debug, Default)]
pub struct StatementStreamReceipt {
    /// Extra grammar/body passes performed by this producer, not original preparation.
    pub producer_body_parses: usize,
    /// Owned normalized statements transferred from original committed preparation.
    pub normalized_handoff_statements: usize,
    pub declaration_header_parses: usize,
    pub reused_declaration_context: bool,
    pub completed_captures: usize,
    pub freshly_lowered_statements: usize,
    pub contextual_fallback_statements: usize,
    pub lowering_failures: usize,
    pub emitted_before_parser_eof: usize,
    pub emitted_statements: usize,
    pub largest_capture_nodes: usize,
    pub parser_committed: bool,
    /// Accepted expression recipes transferred at completed statement boundaries.
    pub expression_recipes: usize,
    pub consumed_expression_recipes: usize,
    /// Committed recipe roots carrying actual Pratt reductions (>0 operators).
    pub source_pratt_expression_recipes: usize,
}

/// Compare exactly the previous canonical Debug byte sequence, without
/// allocating a second AST-sized string. Formatting stops on the first unequal
/// byte; equal prefixes still require the complete expected length.
fn exact_debug_equal<T: std::fmt::Debug>(actual: &T, expected: &T, scratch: &mut String) -> bool {
    use std::fmt::Write;
    scratch.clear();
    if write!(scratch, "{expected:?}").is_err() { return false; }
    struct Compare<'a> { expected: &'a [u8], consumed: usize }
    impl std::fmt::Write for Compare<'_> {
        fn write_str(&mut self, value: &str) -> std::fmt::Result {
            let end = self.consumed.checked_add(value.len()).ok_or(std::fmt::Error)?;
            if self.expected.get(self.consumed..end) != Some(value.as_bytes()) {
                return Err(std::fmt::Error);
            }
            self.consumed = end;
            Ok(())
        }
    }
    let mut compare = Compare { expected: scratch.as_bytes(), consumed: 0 };
    write!(&mut compare, "{actual:?}").is_ok() && compare.consumed == compare.expected.len()
}

struct Producer<'a> {
    source: &'a str,
    walker: PhpWalker,
    ast: Option<direct_ast::Ast<'a>>,
    arena_base: usize,
    comparison_scratch: String,
    statements: Vec<Option<Statement>>,
    next: usize,
    captures: usize,
    poisoned: bool,
    sink: &'a mut dyn FnMut(usize, Statement) -> Result<(), String>,
    receipt: StatementStreamReceipt,
}
impl Producer<'_> {
    fn send_through(&mut self, end: usize, fresh: Option<Statement>, final_flush: bool)
        -> Result<(), BuildError> {
        let mut fresh = fresh;
        while let Some(Some(expected)) = self.statements.get(self.next) {
            let span = expected.span;
            let source_offset = |line: u32, col: u32| if line == 0 { Some(0) } else {
                self.walker.line_starts.get(line as usize - 1)
                    .and_then(|start| start.checked_add(col.saturating_sub(1) as usize))
            };
            let start = source_offset(span.start_line, span.start_col);
            let complete = source_offset(span.end_line, span.end_col);
            // Adjacent statements can begin exactly at this capture's end.
            // Require the complete normalized source range, not only its head,
            // before allowing a parser callback to deliver it.
            if !final_flush && (start.is_none_or(|start| start >= end && start != 0)
                || complete.is_none_or(|complete| complete > end)) { break; }
            // Preserve the exact discovered normalization, including all late
            // class, namespace, slot, alias, bootstrap and module directives.
            // Only an exactly identical fresh node can replace its fallback.
            let identical = fresh.as_ref().is_some_and(|actual|
                exact_debug_equal(actual,expected,&mut self.comparison_scratch));
            let normalized = self.statements[self.next].take().expect("pending normalized statement");
            let statement = if identical {
                self.receipt.freshly_lowered_statements += 1;
                fresh.take().expect("matched fresh statement")
            } else {
                self.receipt.contextual_fallback_statements += 1;
                normalized
            };
            (self.sink)(self.next, statement).map_err(|message| BuildError {
                span: grameo::Span { start: end, end }, message,
            })?;
            self.next += 1;
            self.receipt.emitted_statements += 1;
            if !final_flush && end < self.source.len() {
                self.receipt.emitted_before_parser_eof += 1;
            }
        }
        Ok(())
    }
}
impl Builder for Producer<'_> {
    fn supports_choice_heads(&self) -> bool { self.ast.as_ref().unwrap().supports_choice_heads() }
    fn supports_unobserved_failure_prefixes(&self) -> bool { self.ast.as_ref().unwrap().supports_unobserved_failure_prefixes() }
    fn opaque_rule(&self, rule: usize) -> bool { self.ast.as_ref().unwrap().opaque_rule(rule) }
    fn observes_rule(&self, rule: usize) -> bool {
        rule == Rule::pure_top_level_statement as usize || self.ast.as_ref().unwrap().observes_rule(rule)
    }
    fn observes_capture(&self, rule: usize) -> bool { self.ast.as_ref().unwrap().observes_capture(rule) }
    fn wants_completed_capture(&self, rule: usize) -> bool { rule == Rule::pure_top_level_statement as usize }
    fn begin(&mut self, rule: usize, offset: usize) {
        if self.ast.as_ref().unwrap().observes_rule(rule) { self.ast.as_mut().unwrap().begin(rule, offset); }
    }
    fn finish(&mut self, rule: usize, span: grameo::Span, source: &str) {
        self.try_finish(rule, span, source).expect("use checked statement producer hooks");
    }
    fn try_finish(&mut self, rule: usize, span: grameo::Span, source: &str) -> Result<(),BuildError> {
        if self.ast.as_ref().unwrap().observes_rule(rule) { self.ast.as_mut().unwrap().try_finish(rule, span, source)?; }
        Ok(())
    }
    fn try_finish_capture(&mut self, rule: usize, span: grameo::Span, source: &str,
        capture: Option<grameo::tree::CaptureHandle>) -> Result<(),BuildError> {
        if self.ast.as_ref().unwrap().observes_rule(rule) {
            self.ast.as_mut().unwrap().try_finish_capture(rule, span, source, capture)?;
        }
        Ok(())
    }
    fn supports_pratt_at(&self, rule: usize, offset: usize, source: &str) -> bool {
        self.ast.as_ref().unwrap().supports_pratt_at(rule,offset,source)
    }
    fn accepts_pratt(&self, rule: usize, span: grameo::Span) -> bool { self.ast.as_ref().unwrap().accepts_pratt(rule,span) }
    fn suppress_pratt_captures(&self, rule: usize) -> bool { self.ast.as_ref().unwrap().suppress_pratt_captures(rule) }
    fn reduce(&mut self, rule: usize, fixity: grameo::pratt::Fixity, span: grameo::Span) {
        self.ast.as_mut().unwrap().reduce(rule,fixity,span);
    }
    fn prune_capture_children(&self, rule: usize, span: grameo::Span) -> bool {
        rule == Rule::pure_top_level_statement as usize || self.ast.as_ref().unwrap().prune_capture_children(rule,span)
    }
    fn checkpoint(&self) -> usize { self.arena_base + self.ast.as_ref().unwrap().checkpoint() }
    fn rollback(&mut self, mark: usize) {
        if mark < self.arena_base { self.poisoned = true; }
        else { self.ast.as_mut().unwrap().rollback(mark - self.arena_base); }
    }
    fn rollback_pratt(&mut self, mark: usize, rule: usize, span: grameo::Span) {
        if mark < self.arena_base { self.poisoned = true; }
        else { self.ast.as_mut().unwrap().rollback_pratt(mark - self.arena_base,rule,span); }
    }
    fn try_completed_capture(&mut self, _: usize, span: grameo::Span,
        tree: grameo::ParseTree<'_, '_>) -> Result<(), BuildError> {
        if self.poisoned { return Err(BuildError { span, message: "PHP statement stream rolled back an emitted prefix; abort source unit".into() }); }
        self.captures += 1;
        self.receipt.completed_captures += 1;
        self.receipt.largest_capture_nodes = self.receipt.largest_capture_nodes.max(tree.capture_count());
        let pair = tree.into_typed_pairs::<Rule>().next()
            .and_then(|pair| pair.into_inner().next())
            .ok_or_else(|| BuildError { span, message: "completed PHP statement capture is empty".into() })?;
        // Own a new local arena generation before contextual binding; source
        // scopes stay in the same walker while parser recipe Ids never escape.
        let old = self.ast.take().expect("active statement AST generation");
        let completed_mark = old.checkpoint();
        let (next, expressions, pratt) = old.drain_statement_boundary(span.end);
        self.arena_base = self.arena_base.checked_add(completed_mark).and_then(|mark| mark.checked_add(1))
            .ok_or_else(|| BuildError { span, message: "PHP statement arena generation overflow".into() })?;
        self.ast = Some(next);
        self.receipt.expression_recipes += expressions.len();
        self.receipt.source_pratt_expression_recipes += pratt;
        self.walker.direct_expressions = expressions;
        self.walker.direct_input = Some((self.source.as_ptr() as usize,self.source.len()));
        let fresh = match walk_statement(&mut self.walker, pair) {
            Ok(statement) => statement,
            Err(_) => { self.receipt.lowering_failures += 1; None },
        };
        self.receipt.consumed_expression_recipes += direct_ast::consumed_expression_recipes(&self.walker.direct_expressions);
        self.walker.direct_expressions.clear();
        self.walker.direct_input = None;
        self.send_through(span.end, fresh, false)
    }
    fn try_committed(&mut self, _: &str, consumed: usize) -> Result<(), BuildError> {
        let span = grameo::Span { start: consumed, end: consumed };
        if self.poisoned || consumed != self.source.len() {
            return Err(BuildError { span, message: "PHP statement stream has no valid source-unit commit".into() });
        }
        self.send_through(consumed, None, true)?;
        self.receipt.parser_committed = true;
        Ok(())
    }
}

/// Actual parser producer used by the compiler frontend registry. The semantic
/// manifest must describe this exact source after normalizations and bindings.
/// Callers must abort their unpublished compiler session on any returned error.
/// This opt-in route does not change ordinary PHP/PEG parsing.
pub fn produce_grameo_statements(source: &str, manifest: Module,
    sink: &mut (dyn FnMut(usize, Statement) -> Result<(), String> + '_)) -> Result<(), String> {
    produce_grameo_statements_with_receipt(source, manifest, sink).map(|_| ())
}

/// Validate original source/grammar/normalization only. The owned compiler route
/// deliberately provides no Module or statement replacement opportunity here.
pub(crate) fn validate_prepared_context(source: &str,
    context: &vybe_compiler::statement_codegen::ProducerContext) -> Result<(),String> {
    let (original, commit) = if let Some(owned) = context.downcast_ref::<OwnedSourceContext>() {
        (owned.original_source.as_ref(), owned.committed_preparation)
    } else if let Some(facts) = context.downcast_ref::<PhpDeclarationFacts>() {
        (facts.original_source.as_ref(), facts.committed_preparation)
    } else { return Err("PHP statement context has the wrong type".into()); };
    if source != original { return Err("PHP statement context source mismatch".into()); }
    validate_source_commit(source, commit)?;
    Ok(())
}

pub fn produce_grameo_statements_prepared(source: &str, manifest: Module,
    context: &vybe_compiler::statement_codegen::ProducerContext,
    sink: &mut (dyn FnMut(usize, Statement) -> Result<(), String> + '_)) -> Result<(), String> {
    produce_grameo_statements_prepared_with_receipt(source,manifest,context,sink).map(|_| ())
}
pub fn produce_grameo_statements_prepared_with_receipt(source: &str, manifest: Module,
    context: &vybe_compiler::statement_codegen::ProducerContext,
    sink: &mut (dyn FnMut(usize, Statement) -> Result<(), String> + '_)) -> Result<StatementStreamReceipt,String> {
    validate_prepared_context(source,context)?;
    {
        // This is the ORIGINAL normalized body moved from preparation, including
        // subsequent caller-owned import/entry binding. The compiler sealed that
        // configured manifest before transferring ownership here and still
        // checks each statement receipt/order. No second parse or normalization
        // is necessary to rediscover the same owned statements.
        let mut receipt = StatementStreamReceipt {
            reused_declaration_context: true, parser_committed: true,
            ..StatementStreamReceipt::default()
        };
        for (index, statement) in manifest.body.into_iter().enumerate() {
            sink(index,statement)?;
            receipt.emitted_statements += 1;
            receipt.normalized_handoff_statements += 1;
        }
        Ok(receipt)
    }
}
pub fn produce_grameo_statements_with_receipt(source: &str, manifest: Module,
    sink: &mut (dyn FnMut(usize, Statement) -> Result<(), String> + '_))
    -> Result<StatementStreamReceipt, String> {
    produce_inner(source,manifest,sink,None)
}
fn produce_inner(source: &str, mut manifest: Module,
    sink: &mut (dyn FnMut(usize, Statement) -> Result<(), String> + '_), facts: Option<&PhpDeclarationFacts>)
    -> Result<StatementStreamReceipt,String> {
    let trimmed = source.trim_start();
    let normalize_first = trimmed.starts_with("<?") || (trimmed.starts_with('<') && source.contains("<?"));
    // Ordinary parse also accepts text-first mixed documents through `program`.
    // Validate that existing route before translating it into statement form;
    // a tag-like string in otherwise valid bare PHP is not HTML.
    let mixed_fallback = facts.is_none() && !normalize_first && source.contains("<?")
        && PhpParser::parse(Rule::program_pure, source).is_err();
    if mixed_fallback {
        PhpParser::parse(Rule::program, source)
            .map_err(|error| format!("PHP stream mixed-source parse error: {error}"))?;
    }
    let normalized;
    let source = if let Some(facts) = facts { facts.normalized_source.as_ref() } else if normalize_first || mixed_fallback {
        normalized = normalized_source_for_parse(source); normalized.as_ref()
    } else { source };
    // Discovery retains only scoped declaration facts. It does not lower the
    // body or publish code. The complete semantic manifest is supplied once by
    // the caller; no Module compilation is repeated for individual statements.
    let mut walker = PhpWalker::default();
    walker.line_starts = build_line_starts(source);
    if let Some(facts) = facts { facts.install(&mut walker); }
    else {
        let program = PhpParser::parse(Rule::program_pure, source)
            .map_err(|error| format!("PHP stream declaration parse error: {error}"))?
            .next().ok_or("PHP stream declaration parse is empty")?;
        pre_register_php_type_names(&mut walker, &program);
        pre_register_php_function_signatures(&mut walker, &program);
    }
    let mut builder = Producer { source, walker, ast: Some(direct_ast::Ast::new(source)), arena_base: 0, comparison_scratch: String::new(),
        statements: std::mem::take(&mut manifest.body).into_iter().map(Some).collect(),
        next: 0, captures: 0, poisoned: false, sink,
        receipt: StatementStreamReceipt { producer_body_parses: 1, declaration_header_parses: usize::from(facts.is_none()),
            reused_declaration_context: facts.is_some(), ..StatementStreamReceipt::default() } };
    crate::native_grammar::Parser.build_captures(Rule::program_pure, source, &mut builder)
        .map_err(|error| format!("PHP statement stream parse error: {error}"))?;
    if !builder.receipt.parser_committed || builder.next != builder.statements.len() {
        return Err("PHP statement stream did not deliver its complete normalized manifest".into());
    }
    Ok(builder.receipt)
}
