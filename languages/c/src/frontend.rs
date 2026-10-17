//! Opt-in matching backends at the existing C contextual lowering boundary.
use std::cell::{Cell, RefCell};
use std::collections::HashMap;
use std::rc::Rc;
use std::time::{Duration, Instant};

use crate::{Rule, native_grammar};

#[path = "frontend/owned.rs"]
mod owned;
#[path = "frontend/direct.rs"]
mod direct;
pub(crate) use direct::{BodyItem, InitializerLease, OwnedStatement};
pub(crate) use owned::PlainScalarDeclarator;
pub use owned::declarations::{DeclarationContext, DeclarationPolicy, ScalarPolicy, UnsignedWidth};

pub(crate) fn take_initializer_list(
    pair: &grameo::compat::Pair<'_, Rule>,
) -> Option<vybe_ast::Expression> {
    let value = direct::take_initializer_list(pair);
    if value.is_some() {
        record_construction(|profile| profile.direct_initializer_handoffs += 1);
    }
    value
}

pub(crate) fn lease_initializer_list(
    pair: &grameo::compat::Pair<'_, Rule>,
) -> Option<InitializerLease> {
    direct::lease_initializer_list(pair)
}

pub(crate) fn take_control_statement(
    pair: &grameo::compat::Pair<'_, Rule>,
) -> Option<OwnedStatement> {
    let value = direct::take_control_statement(pair);
    if value.is_some() {
        record_construction(|profile| profile.direct_statement_handoffs += 1);
    }
    value
}

pub(crate) fn take_primary_expression(
    pair: &grameo::compat::Pair<'_, Rule>,
) -> Option<vybe_ast::Expression> {
    let value = direct::take_primary_expression(pair);
    if value.is_some() {
        record_construction(|profile| profile.direct_primary_handoffs += 1);
    }
    value
}

pub(crate) fn take_syntax_expression(
    pair: &grameo::compat::Pair<'_, Rule>,
) -> Option<vybe_ast::Expression> {
    let value = direct::take_syntax_expression(pair);
    if value.is_some() {
        record_construction(|profile| profile.direct_syntax_handoffs += 1);
    }
    value
}

pub(crate) fn take_direct_binary_region<'a>(
    pair: &grameo::compat::Pair<'a, Rule>,
) -> Option<direct::BinaryRegion<'a>> {
    let region = direct::take_binary_region(pair)?;
    record_construction(|profile| {
        profile.direct_binary_handoffs += 1;
        profile.direct_binary_operands += region.operand_count();
    });
    Some(region)
}

pub(crate) fn take_plain_declarator(
    pair: &grameo::compat::Pair<'_, Rule>,
) -> Option<PlainScalarDeclarator> {
    direct::take_plain_declarator(pair)
}

pub(crate) fn record_direct_declaration_handoff() {
    record_construction(|profile| profile.direct_declaration_handoffs += 1);
}

pub(crate) fn take_plain_parameter<'a>(
    pair: &grameo::compat::Pair<'a, Rule>,
) -> Option<direct::PlainParameter<'a>> {
    direct::take_plain_parameter(pair)
}

pub(crate) fn record_direct_parameter_handoff() {
    record_construction(|profile| profile.direct_parameter_handoffs += 1);
}

pub(crate) fn take_direct_if_region<'a>(
    pair: &grameo::compat::Pair<'a, Rule>,
) -> Option<direct::IfRegion<'a>> {
    direct::take_if_region(pair)
}

pub(crate) fn record_direct_if_handoff() {
    record_construction(|profile| profile.direct_if_handoffs += 1);
}

pub(crate) fn take_direct_while_region<'a>(
    pair: &grameo::compat::Pair<'a, Rule>,
) -> Option<direct::WhileRegion<'a>> {
    direct::take_while_region(pair)
}

pub(crate) fn take_direct_do_while_region<'a>(
    pair: &grameo::compat::Pair<'a, Rule>,
) -> Option<direct::DoWhileRegion<'a>> {
    direct::take_do_while_region(pair)
}

pub(crate) fn record_direct_loop_handoff() {
    record_construction(|profile| profile.direct_loop_handoffs += 1);
}

pub(crate) fn take_direct_for_region<'a>(
    pair: &grameo::compat::Pair<'a, Rule>,
) -> Option<direct::ForRegion<'a>> {
    direct::take_for_region(pair)
}

pub(crate) fn take_direct_switch_region<'a>(
    pair: &grameo::compat::Pair<'a, Rule>,
) -> Option<direct::SwitchRegion<'a>> {
    direct::take_switch_region(pair)
}

pub(crate) fn record_direct_switch_handoff() {
    record_construction(|profile| profile.direct_switch_handoffs += 1);
}

pub(crate) use direct::CallRegion;

pub(crate) fn take_direct_call_region<'a>(
    pair: &grameo::compat::Pair<'a, Rule>,
) -> Option<CallRegion<'a>> {
    direct::take_call_region(pair)
}

pub(crate) fn record_direct_call_handoff() {
    record_construction(|profile| profile.direct_call_handoffs += 1);
}

pub(crate) fn take_direct_function_region<'a>(pair: &grameo::compat::Pair<'a, Rule>) -> Option<direct::FunctionRegion<'a>> {
    direct::take_function_region(pair)
}

pub(crate) fn record_direct_function_handoff() {
    record_construction(|profile| profile.direct_function_handoffs += 1);
}

pub(crate) fn take_direct_return_region<'a>(pair: &grameo::compat::Pair<'a, Rule>) -> Option<direct::ReturnRegion<'a>> {
    direct::take_return_region(pair)
}

pub(crate) fn record_direct_return_handoff() {
    record_construction(|profile| profile.direct_return_handoffs += 1);
}

pub(crate) use direct::{IndexRegion, MemberRegion};

pub(crate) fn take_direct_index_region<'a>(pair: &grameo::compat::Pair<'a, Rule>) -> Option<IndexRegion<'a>> {
    direct::take_index_region(pair)
}

pub(crate) fn take_direct_member_region<'a>(pair: &grameo::compat::Pair<'a, Rule>) -> Option<MemberRegion<'a>> {
    direct::take_member_region(pair)
}

pub(crate) fn finish_direct_index(region: Option<IndexRegion<'_>>, object: vybe_ast::Expression,
    index: vybe_ast::Expression) -> vybe_ast::Expression
{
    if let Some(region) = region {
        record_construction(|profile| profile.direct_index_handoffs += 1);
        region.finish(object, index)
    } else {
        vybe_ast::Expression::new(vybe_ast::ExprKind::Index {
            object: Box::new(object), index: Box::new(index), null_safe: false,
        })
    }
}

pub(crate) fn finish_direct_member(region: Option<MemberRegion<'_>>, object: vybe_ast::Expression,
    field: &str) -> vybe_ast::Expression
{
    if let Some(region) = region {
        record_construction(|profile| profile.direct_member_handoffs += 1);
        region.finish(object)
    } else { vybe_platform_libc::emitter::build::member(object, field) }
}

pub(crate) fn take_body<'a>(
    pair: &grameo::compat::Pair<'a, Rule>,
) -> Option<Vec<BodyItem<'a>>> {
    let value = direct::take_body(pair);
    if let Some(items) = &value {
        record_construction(|profile| {
            profile.direct_body_handoffs += 1;
            profile.direct_statement_handoffs += items.iter()
                .filter(|item| matches!(item, BodyItem::Owned { .. })).count();
        });
    }
    value
}

mod optimized_grammar {
    #![allow(dead_code)]
    include!(concat!(env!("OUT_DIR"), "/optimized_parser.rs"));
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum Backend {
    Peg,
    Native,
    Owned,
    /// Retain contextual atoms and journal binary regions without Pratt replay.
    /// Consumers must use `take_binary_with_context` to benefit from the recipes.
    OwnedContextual,
    /// Native matching with raw scalar declarations, then ordered contextual Pratt.
    ContextualPratt,
}

// No C semantic observations: certified failed prefixes may bypass callback
// queries and journals while retaining reference captures and diagnostics.
struct NativeBuilder;

impl grameo::builder::Builder for NativeBuilder {
    fn supports_choice_heads(&self) -> bool { choice_heads_enabled() }
    fn supports_unobserved_failure_prefixes(&self) -> bool { true }
    fn observes_rule(&self, _: usize) -> bool { false }
    fn opaque_rule(&self, _: usize) -> bool { true }
    fn checkpoint(&self) -> usize { 0 }
    fn rollback(&mut self, _: usize) {}
    fn begin(&mut self, _: usize, _: usize) {}
    fn finish(&mut self, _: usize, _: grameo::Span, _: &str) {}
}

thread_local! {
    static BACKEND: Cell<Option<Backend>> = const { Cell::new(None) };
    static OWNED_CACHE: RefCell<Option<SourceCache>> = const { RefCell::new(None) };
    static MIXED_CACHE: RefCell<Option<MixedSourceCache>> = const { RefCell::new(None) };
    static DECLARATION_CACHE: RefCell<Option<DeclarationSourceCache>> = const { RefCell::new(None) };
    static PRATT_CACHE: RefCell<Option<PrattSourceCache>> = const { RefCell::new(None) };
    static SOURCE_PRATT_REQUESTED: Cell<bool> = const { Cell::new(false) };
    static CONSTRUCTION_PROFILE: Cell<Option<ConstructionProfile>> = const { Cell::new(None) };
    static MATCHING_ENABLED: Cell<bool> = const { Cell::new(false) };
    static MATCHING_PROFILE: RefCell<Option<MatchingProfile>> = const { RefCell::new(None) };
    #[cfg(test)]
    static OWNED_HANDOFFS: Cell<usize> = const { Cell::new(0) };
    #[cfg(test)]
    static PRATT_DECLARATIONS_DISABLED: Cell<bool> = const { Cell::new(false) };
    #[cfg(test)]
    static PLAIN_DECLARATOR_SCAN_DISABLED: Cell<bool> = const { Cell::new(false) };
    #[cfg(test)]
    static SOURCE_PRATT_DISABLED: Cell<bool> = const { Cell::new(false) };
    #[cfg(test)]
    static DIRECT_DISABLED: Cell<bool> = const { Cell::new(false) };
    #[cfg(test)]
    static NATIVE_PRATT_ENABLED: Cell<Option<bool>> = const { Cell::new(None) };
    #[cfg(test)]
    static CHOICE_HEADS_DISABLED: Cell<bool> = const { Cell::new(false) };
    #[cfg(test)]
    static GENERAL_UNIT_ENGINE: Cell<bool> = const { Cell::new(false) };
}

fn choice_heads_enabled() -> bool {
    #[cfg(test)]
    return !CHOICE_HEADS_DISABLED.with(Cell::get);
    #[cfg(not(test))]
    true
}

type SourceCache = HashMap<(usize, usize), HashMap<(usize, usize), owned::Cached>>;
type MixedSourceCache = HashMap<(usize, usize), Rc<HashMap<(usize, usize), owned::Mixed>>>;
type DeclarationSourceCache =
    HashMap<(usize, usize), HashMap<(usize, usize), PlainScalarDeclarator>>;
type PrattSourceCache = HashMap<(usize, usize), Rc<HashMap<(usize, usize), owned::PrattRecipe>>>;

#[derive(Clone, Copy, Debug, Default)]
pub struct ConstructionProfile {
    pub parse_calls: usize,
    pub direct_initializer_handoffs: usize,
    pub direct_body_handoffs: usize,
    pub direct_statement_handoffs: usize,
    pub direct_primary_handoffs: usize,
    pub direct_syntax_handoffs: usize,
    pub direct_binary_handoffs: usize,
    pub direct_binary_operands: usize,
    pub direct_declaration_handoffs: usize,
    pub direct_parameter_handoffs: usize,
    pub direct_if_handoffs: usize,
    pub direct_loop_handoffs: usize,
    pub direct_switch_handoffs: usize,
    pub direct_call_handoffs: usize,
    pub direct_index_handoffs: usize,
    pub direct_member_handoffs: usize,
    pub direct_function_handoffs: usize,
    pub direct_return_handoffs: usize,
    pub generated_unit_matches: usize,
    pub general_unit_matches: usize,
    pub scalar_regions: usize,
    pub scalar_handoffs: usize,
    pub mixed_regions: usize,
    pub mixed_handoffs: usize,
    pub contextual_operands: usize,
    pub contextual_pratt_handoffs: usize,
    pub contextual_pratt_operands: usize,
    pub contextual_pratt_single_infix_regions: usize,
    pub contextual_pratt_inline_regions: usize,
    pub contextual_pratt_spilled_regions: usize,
    pub source_pratt_regions: usize,
    pub source_pratt_handoffs: usize,
    pub source_pratt_operands: usize,
    pub source_pratt_matches: usize,
    pub source_pratt_recaptures: usize,
    pub mixed_misses: usize,
    pub declaration_regions: usize,
    pub declaration_handoffs: usize,
    pub declarator_scan_bypasses: usize,
    pub declarator_scan_fallbacks: usize,
}

#[derive(Clone, Copy, Debug, Default)]
pub struct MatchingTotals {
    pub calls: usize,
    pub bytes: usize,
    pub failures: usize,
    pub elapsed: Duration,
}

impl MatchingTotals {
    fn add(&mut self, bytes: usize, elapsed: Duration, failed: bool) {
        self.calls = self.calls.saturating_add(1);
        self.bytes = self.bytes.saturating_add(bytes);
        self.failures = self.failures.saturating_add(usize::from(failed));
        self.elapsed = self.elapsed.saturating_add(elapsed);
    }
}

#[derive(Clone, Debug)]
pub struct RuleMatchingProfile {
    pub backend: Backend,
    pub rule: String,
    pub totals: MatchingTotals,
}

#[derive(Clone, Debug, Default)]
pub struct MatchingProfile {
    pub rules: Vec<RuleMatchingProfile>,
    pub total: MatchingTotals,
    /// Calls beyond the fixed rule/backend bucket limit remain accounted for.
    pub overflow: MatchingTotals,
}

const MATCHING_RULE_LIMIT: usize = 64;

fn record_matching(backend: Backend, rule: &str, bytes: usize, elapsed: Duration, failed: bool) {
    MATCHING_PROFILE.with(|slot| {
        let mut slot = slot.borrow_mut();
        let Some(profile) = slot.as_mut() else { return; };
        profile.total.add(bytes, elapsed, failed);
        if let Some(row) = profile.rules.iter_mut()
            .find(|row| row.backend == backend && row.rule == rule)
        {
            row.totals.add(bytes, elapsed, failed);
        } else if profile.rules.len() < MATCHING_RULE_LIMIT {
            let mut totals = MatchingTotals::default();
            totals.add(bytes, elapsed, failed);
            profile.rules.push(RuleMatchingProfile { backend, rule: rule.into(), totals });
        } else {
            profile.overflow.add(bytes, elapsed, failed);
        }
    });
}

/// Instrument only this thread's capture matching, including builder callbacks,
/// cache publication and any fallback recapture, not subsequent C lowering.
/// Durations are inclusive if matching itself nests. No source/captures are kept;
/// at most 64 rule/backend buckets are retained. Nested scopes and unwinding
/// restore the previous collector. Disabled parsing does not read a clock.
pub fn profile_matching<T>(action: impl FnOnce() -> T) -> (T, MatchingProfile) {
    struct Restore(bool, Option<MatchingProfile>);
    impl Drop for Restore {
        fn drop(&mut self) {
            MATCHING_PROFILE.with(|slot| *slot.borrow_mut() = self.1.take());
            MATCHING_ENABLED.with(|slot| slot.set(self.0));
        }
    }
    let _restore = Restore(
        MATCHING_ENABLED.with(|slot| slot.replace(true)),
        MATCHING_PROFILE.with(|slot| slot.replace(Some(MatchingProfile::default()))),
    );
    let value = action();
    let mut profile = MATCHING_PROFILE.with(|slot| slot.borrow_mut().take().unwrap());
    profile.rules.sort_by_key(|row| std::cmp::Reverse(row.totals.elapsed));
    (value, profile)
}

fn record_construction(change: impl FnOnce(&mut ConstructionProfile)) {
    CONSTRUCTION_PROFILE.with(|slot| {
        if let Some(mut profile) = slot.get() {
            change(&mut profile);
            slot.set(Some(profile));
        }
    });
}

/// Collect construction counts separately from uninstrumented CPU measurements.
/// The calling thread's previous collector is restored on return or unwind.
pub fn profile_construction<T>(action: impl FnOnce() -> T) -> (T, ConstructionProfile) {
    struct Restore(Option<ConstructionProfile>);
    impl Drop for Restore {
        fn drop(&mut self) {
            CONSTRUCTION_PROFILE.with(|slot| slot.set(self.0));
        }
    }
    let _restore = Restore(
        CONSTRUCTION_PROFILE.with(|slot| slot.replace(Some(ConstructionProfile::default()))),
    );
    let value = action();
    let profile = CONSTRUCTION_PROFILE.with(|slot| slot.get().unwrap());
    (value, profile)
}

pub(crate) fn selected_backend(selection: Option<&str>) -> Backend {
    match selection {
        None | Some("native") => Backend::Native,
        Some("owned") => Backend::Owned,
        Some("mixed") => Backend::OwnedContextual,
        Some("pratt") => Backend::ContextualPratt,
        _ => Backend::Peg,
    }
}

pub(crate) fn parse_registered(source: &str) -> Result<vybe_ast::Module, String> {
    if BACKEND.with(Cell::get).is_some() {
        return crate::walker::parse(source);
    }
    let backend = selected_backend(std::env::var("VYBE_C_FRONTEND").ok().as_deref());
    with_backend(backend, || crate::walker::parse(source))
}

/// Select matching for this thread, restoring the previous selection on unwind.
/// Include workers on other threads retain their own backend selection.
pub fn with_backend<T>(backend: Backend, action: impl FnOnce() -> T) -> T {
    struct Restore(
        Option<Backend>,
        Option<SourceCache>,
        Option<MixedSourceCache>,
        Option<DeclarationSourceCache>,
        Option<PrattSourceCache>,
        bool,
    );
    impl Drop for Restore {
        fn drop(&mut self) {
            BACKEND.with(|slot| slot.set(self.0));
            OWNED_CACHE.with(|slot| *slot.borrow_mut() = self.1.take());
            MIXED_CACHE.with(|slot| *slot.borrow_mut() = self.2.take());
            DECLARATION_CACHE.with(|slot| *slot.borrow_mut() = self.3.take());
            PRATT_CACHE.with(|slot| *slot.borrow_mut() = self.4.take());
            SOURCE_PRATT_REQUESTED.with(|slot| slot.set(self.5));
        }
    }
    let source_setting = std::env::var_os("VYBE_C_SOURCE_PRATT_RECEIPTS");
    // Native uses matcher-issued reductions by default. Keep a diagnostic
    // control without changing the registered parser or C lowering policy.
    let source_requested = match backend {
        Backend::Native => !source_setting.as_ref().is_some_and(|value| value == "0"),
        Backend::ContextualPratt => source_setting.as_ref().is_some_and(|value| value == "1"),
        _ => false,
    };
    #[cfg(test)]
    let source_requested = if backend == Backend::Native {
        NATIVE_PRATT_ENABLED.with(Cell::get).unwrap_or(source_requested)
    } else { source_requested };
    let _restore = Restore(
        BACKEND.with(|slot| slot.replace(Some(backend))),
        OWNED_CACHE.with(|slot| slot.replace((backend == Backend::Owned).then(HashMap::new))),
        MIXED_CACHE
            .with(|slot| slot.replace((backend == Backend::OwnedContextual).then(HashMap::new))),
        DECLARATION_CACHE
            .with(|slot| slot.replace(matches!(backend,
                Backend::OwnedContextual | Backend::ContextualPratt).then(HashMap::new))),
        PRATT_CACHE.with(|slot| slot.replace((backend == Backend::ContextualPratt
            || backend == Backend::Native && source_requested).then(HashMap::new))),
        SOURCE_PRATT_REQUESTED.with(|slot| slot.replace(source_requested)),
    );
    direct::with_cache_scope(action)
}

pub(crate) fn parse_captures(
    rule: Rule,
    source: &str,
    options: grameo::MatchOptions,
) -> Result<grameo::compat::Pairs<'_, Rule>, grameo::ParseError> {
    record_construction(|profile| profile.parse_calls += 1);
    let started = MATCHING_ENABLED.with(Cell::get).then(Instant::now);
    let backend = BACKEND.with(Cell::get).unwrap_or(Backend::Peg);
    let result = match backend {
        Backend::Peg => {
            grameo::engine::parse_program(&native_grammar::Parser, rule.name(), source, options)
                .map(|tree| tree.into_typed_pairs())
        }
        Backend::Native => {
            invalidate_pratt_cache(source);
            if SOURCE_PRATT_REQUESTED.with(Cell::get) && source_pratt_entry(rule)
                && source_pratt_enabled(options) {
                parse_source_pratt_captures(rule, source, options, false)
            } else if direct_ast_enabled() {
                direct::build_captures(source, true, |builder| {
                    build_hybrid(rule, source, options, builder)
                        .map(|tree| tree.into_typed_pairs())
                })
            } else {
                build_hybrid(rule, source, options, &mut NativeBuilder)
                    .map(|tree| tree.into_typed_pairs())
            }
        }
        Backend::Owned => parse_owned_captures(rule, source, options),
        Backend::OwnedContextual | Backend::ContextualPratt => parse_contextual_captures(rule, source, options),
    };
    if let Some(started) = started {
        record_matching(backend, rule.name(), source.len(), started.elapsed(), result.is_err());
    }
    result
}

/// Match independent capture arenas, retaining the whole-source declaration
/// barrier: callers receive no units until the complete sequence validates.
pub(crate) fn parse_top_level_units(
    source: &str,
    options: grameo::MatchOptions,
) -> Result<Vec<grameo::compat::Pair<'_, Rule>>, grameo::ParseError> {
    let whole_program = || parse_captures(Rule::program, source, options)
        .map(|mut pairs| pairs.next().expect("program anchor").into_inner()
            .filter(|pair| pair.as_rule() != Rule::EOI).collect());
    if BACKEND.with(Cell::get) != Some(Backend::Native)
        || !SOURCE_PRATT_REQUESTED.with(Cell::get) || !source_pratt_enabled(options) {
        return whole_program();
    }
    invalidate_pratt_cache(source);
    direct::reset_unit_captures(source);
    let mut items = Vec::new();
    let mut recipes = HashMap::new();
    let mut offset = 0;
    loop {
        let mut builder = owned::SourcePrattBuilder::native(source);
        record_construction(|profile| {
            profile.parse_calls += 1;
            profile.source_pratt_matches += 1;
        });
        let started = MATCHING_ENABLED.with(Cell::get).then(Instant::now);
        let matched = if direct_ast_enabled() {
            direct::build_unit_captures(source, true, |direct| {
                build_hybrid_unit_inner(Rule::frontend_unit, source, offset, options,
                    &mut WithDirect::new(&mut builder, direct))
            })
        } else {
            build_hybrid_unit_inner(Rule::frontend_unit, source, offset, options, &mut builder)
        };
        if let Some(started) = started {
            record_matching(Backend::Native, Rule::frontend_unit.name(), source.len() - offset,
                started.elapsed(), matched.is_err());
        }
        if let Ok(tree) = matched {
            let end = tree.consumed();
            if end > offset {
                record_construction(|profile| profile.source_pratt_regions += builder.expressions.len());
                recipes.extend(builder.expressions);
                let unit = tree.into_typed_pairs::<Rule>().next().expect("unit anchor");
                items.extend(unit.into_inner().filter(|pair|
                    !matches!(pair.as_rule(), Rule::frontend_trivia | Rule::EOI)));
                offset = end;
                continue;
            }
        } else if let Ok(tree) = grameo::engine::parse_program_at(&native_grammar::Parser,
            Rule::frontend_end.name(), source, offset, options)
            && tree.consumed() == source.len() {
            break;
        }
        // A late failure must discard every earlier handoff. Whole-program
        // recapture preserves original diagnostics and resource accounting.
        invalidate_pratt_cache(source);
        direct::reset_unit_captures(source);
        record_construction(|profile| profile.source_pratt_recaptures += 1);
        return whole_program();
    }
    PRATT_CACHE.with(|slot| {
        if let Some(cache) = slot.borrow_mut().as_mut() {
            if !recipes.is_empty() {
                cache.insert((source.as_ptr() as usize, source.len()), Rc::new(recipes));
            }
        }
    });
    Ok(items)
}

/// Build from validated captures in the caller's current C scope. Matching and
/// whole-module preparation are unchanged; this is not streaming compilation.
pub(crate) fn construct_contextual_binary<'a>(
    pair: &grameo::compat::Pair<'a, Rule>,
    mut operand: impl FnMut(grameo::compat::Pair<'a, Rule>) -> vybe_ast::Expression,
) -> Option<vybe_ast::Expression> {
    let backend = BACKEND.with(Cell::get);
    if pair.as_rule() != Rule::binary_expression
        || backend != Some(Backend::ContextualPratt)
            && !(backend == Some(Backend::Native) && SOURCE_PRATT_REQUESTED.with(Cell::get)) {
        return None;
    }
    let span = pair.as_span();
    let source = span.get_input();
    // Release the TLS borrow before invoking C callbacks: nested expressions
    // and reparses may consult or replace the current source's receipt map.
    let recipes = PRATT_CACHE.with(|slot| slot.borrow().as_ref()?
        .get(&(source.as_ptr() as usize, source.len())).cloned());
    if let Some(recipe) = recipes.as_ref().and_then(|recipes| recipes.get(&(span.start(), span.end())))
        && let Some(value) = recipe.lower(pair, &mut operand) {
        record_construction(|profile| {
            profile.source_pratt_handoffs += 1;
            profile.source_pratt_operands += recipe.operand_count();
            profile.contextual_pratt_handoffs += 1;
            profile.contextual_pratt_operands += recipe.operand_count();
        });
        return Some(value);
    }
    // Native keeps its original typed fold whenever this exact capture lacks a
    // committed receipt. It does not acquire the experimental declaration path.
    if backend == Some(Backend::Native) { return None; }
    let value = owned::construct_contextual_binary(pair, operand);
    if value.is_some() {
        record_construction(|profile| {
            profile.contextual_pratt_handoffs += 1;
            let operands = pair.clone().into_inner()
                .filter(|child| child.as_rule() == Rule::unary_expression).count();
            profile.contextual_pratt_operands += operands;
            if operands == 2 { profile.contextual_pratt_single_infix_regions += 1; }
            if operands <= 4 { profile.contextual_pratt_inline_regions += 1; }
            else { profile.contextual_pratt_spilled_regions += 1; }
        });
    }
    value
}

fn build_hybrid<'a, B: grameo::builder::Builder>(
    rule: Rule,
    source: &'a str,
    options: grameo::MatchOptions,
    builder: &mut B,
) -> Result<grameo::ParseTree<'static, 'a>, grameo::ParseError> {
    // Native/default activation is independent. In the explicit Pratt lane,
    // positional initializer ASTs and expression receipts share this one match.
    if BACKEND.with(Cell::get) == Some(Backend::ContextualPratt) && direct_ast_enabled() {
        return direct::build_captures(source, true, |direct| {
            build_hybrid_inner(rule, source, options, &mut WithDirect::new(builder, direct))
        });
    }
    build_hybrid_inner(rule, source, options, builder)
}

struct WithDirect<'b, 's, B> {
    primary: &'b mut B,
    direct: &'b mut direct::Direct<'s>,
    frames: grameo::arena::EditArena<(usize, bool, bool), (), ()>,
    marks: Vec<(usize, usize, usize)>,
}

impl<'b, 's, B: grameo::builder::Builder> WithDirect<'b, 's, B> {
    fn new(primary: &'b mut B, direct: &'b mut direct::Direct<'s>) -> Self {
        let mut value = Self { primary, direct, frames: Default::default(), marks: Vec::new() };
        value.save();
        value
    }
    fn save(&mut self) {
        use grameo::builder::Builder;
        let mark = (self.primary.checkpoint(), self.direct.checkpoint(), self.frames.checkpoint());
        if self.marks.last() != Some(&mark) { self.marks.push(mark); }
    }
}

impl<B: grameo::builder::Builder> grameo::builder::Builder for WithDirect<'_, '_, B> {
    fn supports_choice_heads(&self) -> bool {
        self.primary.supports_choice_heads() && self.direct.supports_choice_heads()
    }
    fn supports_unobserved_failure_prefixes(&self) -> bool {
        self.primary.supports_unobserved_failure_prefixes()
            && self.direct.supports_unobserved_failure_prefixes()
    }
    fn observes_rule(&self, rule: usize) -> bool {
        self.primary.observes_rule(rule) || self.direct.observes_rule(rule)
    }
    fn observes_capture(&self, rule: usize) -> bool {
        self.primary.observes_capture(rule) || self.direct.observes_capture(rule)
    }
    fn opaque_rule(&self, rule: usize) -> bool {
        // Direct starts observations at non-atomic initializer_list boundaries.
        // Outside those regions, unobserved atomic lexical rules remain inert.
        self.primary.opaque_rule(rule)
            && (!self.direct.observes_rule(rule) || self.direct.opaque_rule(rule))
    }
    fn supports_pratt_at(&self, rule: usize, offset: usize, source: &str) -> bool {
        self.primary.supports_pratt_at(rule, offset, source)
    }
    fn supports_pratt(&self, rule: usize) -> bool { self.primary.supports_pratt(rule) }
    fn accepts_pratt(&self, rule: usize, span: grameo::Span) -> bool {
        self.primary.accepts_pratt(rule, span)
    }
    fn supports_owned_capture_boundary(&self, rule: usize) -> bool {
        // The engine independently rejects elision if the primary observes any
        // omitted capture. Source Pratt's unary handles therefore stay intact.
        self.direct.supports_owned_capture_boundary(rule)
    }
    fn owned_capture_end(&self, rule: usize, offset: usize, source: &str) -> Option<usize> {
        self.direct.owned_capture_end(rule, offset, source)
    }
    fn checkpoint(&self) -> usize { self.marks.len() - 1 }
    fn rollback(&mut self, mark: usize) {
        let (primary, direct, frames) = self.marks[mark];
        self.primary.rollback(primary);
        self.direct.rollback(direct);
        self.frames.rollback(frames);
        self.marks.truncate(mark + 1);
    }
    fn begin(&mut self, _: usize, _: usize) { unreachable!("checked composed begin required") }
    fn try_begin(&mut self, rule: usize, offset: usize) -> Result<(), grameo::builder::BuildError> {
        let primary = self.primary.observes_rule(rule);
        let direct = self.direct.observes_rule(rule);
        let id = self.frames.alloc((rule, primary, direct));
        self.frames.push(id);
        if primary { self.primary.try_begin(rule, offset)?; }
        if direct { self.direct.try_begin(rule, offset)?; }
        self.save();
        Ok(())
    }
    fn finish(&mut self, _: usize, _: grameo::Span, _: &str) {
        unreachable!("checked composed finish required")
    }
    fn try_finish_capture(&mut self, rule: usize, span: grameo::Span, source: &str,
        capture: Option<grameo::tree::CaptureHandle>,
    ) -> Result<(), grameo::builder::BuildError> {
        let id = self.frames.pop().unwrap();
        let (owner, primary, direct) = *self.frames.get(id);
        assert_eq!(owner, rule);
        if primary { self.primary.try_finish_capture(rule, span, source, capture)?; }
        if direct { self.direct.try_finish_capture(rule, span, source, capture)?; }
        self.save();
        Ok(())
    }
    fn try_reduce(&mut self, rule: usize, fixity: grameo::pratt::Fixity, span: grameo::Span)
        -> Result<(), grameo::builder::BuildError> {
        self.primary.try_reduce(rule, fixity, span)?;
        self.direct.try_reduce(rule, fixity, span)?;
        self.save();
        Ok(())
    }
    fn try_pratt_boundary(&mut self, rule: usize, offset: usize)
        -> Result<(), grameo::builder::BuildError> {
        self.primary.try_pratt_boundary(rule, offset)?;
        self.direct.try_pratt_boundary(rule, offset)?;
        self.save();
        Ok(())
    }
    fn committed(&mut self, source: &str, end: usize) {
        assert!(self.frames.values().is_empty());
        self.primary.committed(source, end);
        self.direct.committed(source, end);
    }
}

fn build_hybrid_unit_inner<'a, B: grameo::builder::Builder>(
    rule: Rule,
    source: &'a str,
    offset: usize,
    options: grameo::MatchOptions,
    builder: &mut B,
) -> Result<grameo::ParseTree<'static, 'a>, grameo::ParseError> {
    let grameo::MatchOptions { max_steps, max_rule_depth } = options;
    let generated = max_steps == usize::MAX
        && max_rule_depth == grameo::MatchOptions::default().max_rule_depth;
    #[cfg(test)]
    let generated = generated && !GENERAL_UNIT_ENGINE.with(Cell::get);
    if generated {
        record_construction(|profile| profile.generated_unit_matches += 1);
        let native_rule = optimized_grammar::Rule::from_capture(
            grameo::CaptureRule::User(rule as usize))
            .expect("both generated parsers use the same grammar rule order");
        optimized_grammar::Parser.build_captures_at(native_rule, source, offset, builder)
    } else {
        record_construction(|profile| profile.general_unit_matches += 1);
        grameo::engine::build_capture_program_at(&optimized_grammar::Parser,
            rule.name(), source, offset, options, builder)
    }
}

#[cfg(test)]
#[path = "unit_engine_tests.rs"]
mod unit_engine_tests;

fn build_hybrid_inner<'a, B: grameo::builder::Builder>(
    rule: Rule,
    source: &'a str,
    options: grameo::MatchOptions,
    builder: &mut B,
) -> Result<grameo::ParseTree<'static, 'a>, grameo::ParseError> {
    let parser = &optimized_grammar::Parser;
    if options.max_steps == usize::MAX
        && options.max_rule_depth == grameo::MatchOptions::default().max_rule_depth
    {
        let native_rule =
            optimized_grammar::Rule::from_capture(grameo::CaptureRule::User(rule as usize))
                .expect("both generated parsers use the same grammar rule order");
        parser.build_captures(native_rule, source, builder)
    } else {
        grameo::engine::build_capture_program(parser, rule.name(), source, options, builder)
    }
}

/// Move an owned raw expression or replay it before the existing typed lowering.
pub(crate) fn take_binary(pair: &grameo::compat::Pair<'_, Rule>) -> Option<vybe_ast::Expression> {
    if BACKEND.with(Cell::get) != Some(Backend::Owned) {
        return None;
    }
    let span = pair.as_span();
    let source = span.get_input();
    OWNED_CACHE.with(|cache| {
        cache
            .borrow_mut()
            .as_mut()?
            .get_mut(&(source.as_ptr() as usize, source.len()))?
            .get_mut(&(span.start(), span.end()))
            .map(|value| {
                record_construction(|profile| profile.scalar_handoffs += 1);
                #[cfg(test)]
                OWNED_HANDOFFS.with(|count| count.set(count.get() + 1));
                value.take(source)
            })
    })
}

/// Resolve each original unary capture in the current C scope, then move the
/// resulting operands into a binary AST. The caller retains the original
/// per-operand casts and all post-binary normalization. Repeated reads resolve
/// again; no scope-bound AST is cached under a byte span.
pub(crate) fn take_binary_with_context<'a>(
    pair: &grameo::compat::Pair<'a, Rule>,
    operand: impl FnMut(grameo::compat::Pair<'a, Rule>) -> vybe_ast::Expression,
) -> Option<vybe_ast::Expression> {
    if pair.as_rule() != Rule::binary_expression {
        return None;
    }
    if BACKEND.with(Cell::get) != Some(Backend::OwnedContextual) {
        return None;
    }
    let span = pair.as_span();
    let source = span.get_input();
    // One shared immutable map keeps nested reads and reparses safe without
    // removing/reinserting entries or holding a RefCell borrow in the context.
    let regions = MIXED_CACHE.with(|cache| {
        cache
            .borrow()
            .as_ref()?
            .get(&(source.as_ptr() as usize, source.len()))
            .cloned()
    });
    let recipe = regions
        .as_ref()
        .and_then(|regions| regions.get(&(span.start(), span.end())));
    let result = recipe.and_then(|recipe| recipe.lower(pair, operand));
    record_construction(|profile| {
        if result.is_some() {
            profile.mixed_handoffs += 1;
            profile.contextual_operands += recipe.unwrap().operand_count();
        } else {
            profile.mixed_misses += 1;
        }
    });
    result
}

fn invalidate_pratt_cache(source: &str) {
    let key = (source.as_ptr() as usize, source.len());
    PRATT_CACHE.with(|cache| {
        if let Some(cache) = cache.borrow_mut().as_mut() { cache.remove(&key); }
    });
}

fn source_pratt_entry(rule: Rule) -> bool {
    matches!(rule, Rule::program | Rule::expression | Rule::assignment_expression
        | Rule::binary_expression | Rule::unary_expression | Rule::postfix_expression
        | Rule::compound_statement)
}

fn build_receipt_hybrid<'a, B: grameo::builder::Builder>(
    rule: Rule,
    source: &'a str,
    options: grameo::MatchOptions,
    builder: &mut B,
) -> Result<grameo::ParseTree<'static, 'a>, grameo::ParseError> {
    if BACKEND.with(Cell::get) == Some(Backend::Native) && direct_ast_enabled() {
        return direct::build_captures(source, true, |direct| {
            build_hybrid_inner(rule, source, options, &mut WithDirect::new(builder, direct))
        });
    }
    build_hybrid(rule, source, options, builder)
}

fn parse_source_pratt_captures(
    rule: Rule,
    source: &str,
    options: grameo::MatchOptions,
    declaration_hooks: bool,
) -> Result<grameo::compat::Pairs<'_, Rule>, grameo::ParseError> {
    let key = (source.as_ptr() as usize, source.len());
    let mut builder = if declaration_hooks {
        owned::SourcePrattBuilder::new(source)
    } else {
        owned::SourcePrattBuilder::native(source)
    };
    record_construction(|profile| profile.source_pratt_matches += 1);
    let tree = match build_receipt_hybrid(rule, source, options, &mut builder) {
        Ok(tree) => tree,
        Err(_) => {
            // Preserve the original PEG diagnostics and resource accounting.
            record_construction(|profile| profile.source_pratt_recaptures += 1);
            return build_receipt_hybrid(rule, source, options, &mut NativeBuilder)
                .map(|tree| tree.into_typed_pairs());
        }
    };
    record_construction(|profile| profile.source_pratt_regions += builder.expressions.len());
    let recipes = std::mem::take(&mut builder.expressions);
    let declarations = builder.into_declarations();
    record_construction(|profile| profile.declaration_regions += declarations.len());
    PRATT_CACHE.with(|slot| {
        if let Some(cache) = slot.borrow_mut().as_mut() {
            if !recipes.is_empty() { cache.insert(key, Rc::new(recipes)); }
        }
    });
    if declaration_hooks {
        DECLARATION_CACHE.with(|slot| {
            if let Some(cache) = slot.borrow_mut().as_mut() { cache.insert(key, declarations); }
        });
    }
    Ok(tree.into_typed_pairs())
}

fn parse_contextual_captures(
    rule: Rule,
    source: &str,
    options: grameo::MatchOptions,
) -> Result<grameo::compat::Pairs<'_, Rule>, grameo::ParseError> {
    let key = (source.as_ptr() as usize, source.len());
    invalidate_pratt_cache(source);
    MIXED_CACHE.with(|cache| {
        if let Some(cache) = cache.borrow_mut().as_mut() {
            cache.remove(&key);
        }
    });
    DECLARATION_CACHE.with(|cache| {
        if let Some(cache) = cache.borrow_mut().as_mut() {
            cache.remove(&key);
        }
    });
    #[cfg(test)]
    if BACKEND.with(Cell::get) == Some(Backend::ContextualPratt)
        && PRATT_DECLARATIONS_DISABLED.with(Cell::get)
    {
        return build_hybrid(rule, source, options, &mut NativeBuilder)
            .map(|tree| tree.into_typed_pairs());
    }
    // Identifier probes cannot contain declaration handoffs. Retain the native
    // no-observer path, after invalidating any previous entry for this source.
    if rule == Rule::ident_name {
        return build_hybrid(rule, source, options, &mut NativeBuilder)
            .map(|tree| tree.into_typed_pairs());
    }
    if BACKEND.with(Cell::get) == Some(Backend::ContextualPratt)
        && source_pratt_entry(rule)
        && source_pratt_enabled(options) {
        return parse_source_pratt_captures(rule, source, options, true);
    }
    let mut builder = if BACKEND.with(Cell::get) == Some(Backend::ContextualPratt) {
        owned::ContextualBuilder::declarations_only(source)
    } else {
        owned::ContextualBuilder::new(source)
    };
    let tree = build_hybrid(rule, source, options, &mut builder)?;
    record_construction(|profile| {
        profile.mixed_regions += builder.mixed.expressions.len();
        profile.declaration_regions += builder.declarations.len();
    });
    MIXED_CACHE.with(|cache| {
        if let Some(cache) = cache.borrow_mut().as_mut() {
            cache.insert(key, Rc::new(builder.mixed.expressions));
        }
    });
    DECLARATION_CACHE.with(|cache| {
        if let Some(cache) = cache.borrow_mut().as_mut() {
            cache.insert(key, builder.declarations);
        }
    });
    Ok(tree.into_typed_pairs())
}

fn source_pratt_enabled(options: grameo::MatchOptions) -> bool {
    #[cfg(test)]
    if SOURCE_PRATT_DISABLED.with(Cell::get) { return false; }
    #[cfg(not(test))]
    if !SOURCE_PRATT_REQUESTED.with(Cell::get) { return false; }
    if options.max_steps != usize::MAX
        || options.max_rule_depth != grameo::MatchOptions::default().max_rule_depth {
        return false;
    }
    static CONTRACT: std::sync::OnceLock<bool> = std::sync::OnceLock::new();
    *CONTRACT.get_or_init(|| owned::source_pratt_capture_contract(&optimized_grammar::Parser))
}

fn direct_ast_enabled() -> bool {
    #[cfg(test)]
    return !DIRECT_DISABLED.with(Cell::get);
    #[cfg(not(test))]
    true
}

pub(crate) fn take_scalar_declarator(
    pair: &grameo::compat::Pair<'_, Rule>,
) -> Option<PlainScalarDeclarator> {
    if !matches!(BACKEND.with(Cell::get),
        Some(Backend::OwnedContextual | Backend::ContextualPratt)) {
        return None;
    }
    let span = pair.as_span();
    let source = span.get_input();
    let declaration = DECLARATION_CACHE.with(|cache| {
        cache
            .borrow_mut()
            .as_mut()?
            .get_mut(&(source.as_ptr() as usize, source.len()))?
            .remove(&(span.start(), span.end()))
    });
    if declaration.is_some() {
        record_construction(|p| p.declaration_handoffs += 1);
    }
    declaration
}

pub(crate) fn record_declarator_scan(bypassed: bool) {
    record_construction(|profile| {
        if bypassed {
            profile.declarator_scan_bypasses += 1;
        } else {
            profile.declarator_scan_fallbacks += 1;
        }
    });
}

pub(crate) fn plain_declarator_scan_enabled() -> bool {
    #[cfg(test)]
    return !PLAIN_DECLARATOR_SCAN_DISABLED.with(Cell::get);
    #[cfg(not(test))]
    true
}

/// Stage a raw scalar declaration using already resolved C policies. `None` means
/// that its grammar shape still needs contextual declaration lowering. This
/// does not normalize initializers, register names, update scopes, or publish
/// compiler work; those remain at the ordered contextual declaration boundary.
pub fn construct_scalar_declaration<'a>(
    source: &'a str,
    context: DeclarationContext<'a>,
    options: grameo::MatchOptions,
) -> Result<Option<vybe_ast::Statement>, grameo::ParseError> {
    use grameo::ast::builder::AstBuilder;
    use owned::declarations::{BINDINGS, DeclarationTarget, Value};
    let mut builder = AstBuilder::<DeclarationTarget>::with_context(BINDINGS, context);
    build_hybrid(Rule::normal_declaration, source, options, &mut builder)?;
    Ok(
        match builder
            .into_value()
            .expect("normal declaration constructs one staged value")
        {
            Value::Declaration { statement, .. } => Some(statement),
            _ => None,
        },
    )
}

fn parse_owned_captures(
    rule: Rule,
    source: &str,
    options: grameo::MatchOptions,
) -> Result<grameo::compat::Pairs<'_, Rule>, grameo::ParseError> {
    let key = (source.as_ptr() as usize, source.len());
    OWNED_CACHE.with(|cache| {
        if let Some(cache) = cache.borrow_mut().as_mut() {
            cache.remove(&key);
        }
    });
    // An identifier-only entry cannot construct an owned expression. Keep its
    // captures, but avoid semantic builder work and empty cache publication.
    if rule == Rule::ident_name {
        return build_hybrid(rule, source, options, &mut NativeBuilder)
            .map(|tree| tree.into_typed_pairs());
    }
    let mut builder = owned::Owned::new(source);
    let tree = match build_hybrid(rule, source, options, &mut builder) {
        Ok(tree) => tree,
        Err(error) => {
            // A whole-parse recapture must not publish speculative owned nodes.
            return recapture_owned_syntax(rule, source, options, error)
                .map(|tree| tree.into_typed_pairs());
        }
    };
    record_construction(|profile| profile.scalar_regions += builder.expressions.len());
    OWNED_CACHE.with(|cache| {
        if let Some(cache) = cache.borrow_mut().as_mut() {
            cache.insert(key, builder.expressions);
        }
    });
    Ok(tree.into_typed_pairs())
}

fn recapture_owned_syntax<'a>(
    rule: Rule,
    source: &'a str,
    options: grameo::MatchOptions,
    error: grameo::ParseError,
) -> Result<grameo::ParseTree<'static, 'a>, grameo::ParseError> {
    if error.kind == grameo::ParseErrorKind::Syntax {
        grameo::engine::parse_program(&native_grammar::Parser, rule.name(), source, options)
    } else {
        Err(error)
    }
}

#[cfg(test)]
#[path = "../tests/support/frontend_unit_compatibility.rs"]
mod unit_compatibility;

#[cfg(test)]
mod tests {
    use super::*;

    fn with_choice_heads<T>(enabled: bool, action: impl FnOnce() -> T) -> T {
        struct Restore(bool);
        impl Drop for Restore {
            fn drop(&mut self) { CHOICE_HEADS_DISABLED.with(|slot| slot.set(self.0)); }
        }
        let _restore = Restore(CHOICE_HEADS_DISABLED.with(|slot| slot.replace(!enabled)));
        action()
    }

    #[test]
    fn native_choice_heads_have_metadata_and_composed_capability() {
        use grameo::{builder::Builder, program::{Instruction, Program}};
        let grammar = optimized_grammar::Parser;
        assert!(grammar.choice_heads_metadata_is_current());
        for rule in [Rule::top_level_item, Rule::statement,
            Rule::type_specifier_strict, Rule::type_specifier] {
            let mut root = grammar.rule(rule as usize).expression;
            while let Instruction::Group(child) = grammar.instruction(root) { root = child; }
            assert!(grammar.choice_heads_candidate(root), "missing choice metadata: {rule:?}");
            assert!(grammar.choice_heads(root).is_some(), "missing choice heads: {rule:?}");
        }
        for enabled in [false, true] {
            with_choice_heads(enabled, || {
                let mut native = NativeBuilder;
                let mut source = owned::SourcePrattBuilder::native("");
                let mut direct = direct::Direct::new("");
                assert_eq!(native.supports_choice_heads(), enabled);
                assert_eq!(source.supports_choice_heads(), enabled);
                assert_eq!(direct.supports_choice_heads(), enabled);
                assert_eq!(WithDirect::new(&mut native, &mut direct).supports_choice_heads(), enabled);
                assert_eq!(WithDirect::new(&mut source, &mut direct).supports_choice_heads(), enabled);
                let mut unlicensed = owned::Owned::new("");
                assert!(!WithDirect::new(&mut unlicensed, &mut direct).supports_choice_heads());
            });
        }
    }

    #[test]
    fn native_choice_heads_preserve_complete_modules_captures_and_budgets() {
        for source in [
            "/* lead */ typedef unsigned short T; static T a[]={1,2}; int f(int x){return a[x]+x*2;}",
            "struct P{int x;}; int f(struct P *p){if(p)return p->x; else return 0;}",
            "enum E{A=1,B}; int f(int n){while(n){if(n==2)break; --n;} return n;}",
            "typedef int (*F)(int,...); int f(F cb){return cb(2,3)+1;}",
            "int f(void){int iffy=2; return sizeof(int)+(iffy);}",
            "int f(void){return 1+;}",
            "int f(void){if(1) return 2; else }",
        ] {
            let expected = crate::parse_with_peg(source).map(|value| format!("{value:?}"));
            for enabled in [false, true] {
                let actual = with_native_pratt(true, || with_choice_heads(enabled,
                    || crate::parse_optimized(source))).map(|value| format!("{value:?}"));
                assert_eq!(actual, expected, "choice_heads={enabled}: {source}");
            }
            for max_steps in [0, 1, 50, 512, usize::MAX] {
                for max_rule_depth in [1, 8, grameo::MatchOptions::default().max_rule_depth] {
                    let options = grameo::MatchOptions { max_steps, max_rule_depth };
                    let run = |enabled| with_native_pratt(true, || with_choice_heads(enabled,
                        || with_backend(Backend::Native, ||
                            parse_captures(Rule::program, source, options).map(|pairs|
                                pairs.flatten().map(|pair| (pair.as_rule(), pair.as_span().start(),
                                    pair.as_span().end(), pair.as_str().to_string())).collect::<Vec<_>>())
                                .map_err(|error| format!("{error:?}")))));
                    assert_eq!(run(true), run(false), "{options:?}: {source}");
                }
            }
        }
    }

    fn with_source_pratt<T>(enabled: bool, action: impl FnOnce() -> T) -> T {
        struct Restore(bool);
        impl Drop for Restore {
            fn drop(&mut self) { SOURCE_PRATT_DISABLED.with(|slot| slot.set(self.0)); }
        }
        let _restore = Restore(SOURCE_PRATT_DISABLED.with(|slot| slot.replace(!enabled)));
        action()
    }

    fn with_direct_ast<T>(enabled: bool, action: impl FnOnce() -> T) -> T {
        struct Restore(bool);
        impl Drop for Restore {
            fn drop(&mut self) { DIRECT_DISABLED.with(|slot| slot.set(self.0)); }
        }
        let _restore = Restore(DIRECT_DISABLED.with(|slot| slot.replace(!enabled)));
        action()
    }

    fn with_native_pratt<T>(enabled: bool, action: impl FnOnce() -> T) -> T {
        struct Restore(Option<bool>);
        impl Drop for Restore {
            fn drop(&mut self) { NATIVE_PRATT_ENABLED.with(|slot| slot.set(self.0)); }
        }
        let _restore = Restore(NATIVE_PRATT_ENABLED.with(|slot| slot.replace(Some(enabled))));
        with_source_pratt(enabled, action)
    }

    #[test]
    fn native_pratt_uses_contextual_lowering_without_declaration_journal() {
        for source in [
            "#define K 3\ntypedef short T; int f(T *p,int n){int a[n]; return p[n-1]+sizeof(a)*K;}",
            "int a[]={1,2,3}; int f(int x){return a[x]+x*2;}",
            "int f(int (*cb)(int,...),int x){return cb(x,x+1)+cb(x+2,x);}",
            "int f(int *p,int *q){return (p-q)+((p<q)&&(p!=0));}",
            "typedef int T; int f(int x){return sizeof(T)+({int T=3; T+x;})*x;}",
            "double f(double a,double b){double c=a+b; return c/b+b*2.0-1.0;}",
            "int f(int x){return ++x+x++*x;}",
            "int f(int x){return x+;}",
            "int f(int x){return x*(2+);}",
        ] {
            let reference = crate::parse_with_peg(source).map(|module| format!("{module:?}"));
            for direct in [false, true] {
                let (actual, counts) = profile_construction(|| with_direct_ast(direct,
                    || with_native_pratt(true, || crate::parse_optimized(source))));
                assert_eq!(actual.map(|module| format!("{module:?}")), reference, "direct={direct}: {source}");
                assert_eq!(counts.declaration_regions, 0, "{counts:?}");
                assert_eq!(counts.declaration_handoffs, 0, "{counts:?}");
                if reference.is_ok() {
                    assert!(counts.source_pratt_handoffs + counts.direct_binary_handoffs > 0,
                        "{counts:?}: {source}");
                }
            }
        }
    }

    #[test]
    fn native_pratt_all_operator_pairs_match_complete_ast() {
        let operators = ["||", "&&", "|", "^", "&", "==", "!=", "<", "<=", ">", ">=",
            "<<", ">>", "+", "-", "*", "/", "%"];
        for left in operators {
            for right in operators {
                let source = format!("int f(int a,int b,int c){{return a {left} b {right} c;}}");
                let reference = crate::parse_with_peg(&source).unwrap();
                let (actual, counts) = profile_construction(|| with_native_pratt(true,
                    || crate::parse_optimized(&source)));
                assert_eq!(format!("{:?}", actual.unwrap()), format!("{reference:?}"), "{source}");
                assert!(counts.source_pratt_handoffs + counts.direct_binary_handoffs > 0,
                    "{counts:?}");
                assert_eq!(counts.declaration_handoffs, 0);
            }
        }
    }

    #[test]
    fn native_pratt_preserves_capture_shapes_and_original_resource_errors() {
        for (rule, source) in [
            (Rule::binary_expression, "a+b*c-d"),
            (Rule::binary_expression, "a+b*"),
            (Rule::unary_expression, "sizeof(int[n+1])"),
            (Rule::program, "int f(int a){return a+2*3;}"),
            (Rule::program, "int f(int a){return a+;}"),
            (Rule::program, "int f(int a){return a+2;} int broken( {"),
        ] {
            for max_steps in (0..128).chain([512, 4096, usize::MAX]) {
                for max_rule_depth in [0, 1, grameo::MatchOptions::default().max_rule_depth] {
                    let options = grameo::MatchOptions { max_steps, max_rule_depth };
                    let run = |enabled| with_direct_ast(false, || with_native_pratt(enabled,
                        || with_backend(Backend::Native, || parse_captures(rule, source, options)
                            .map(|pairs| pairs.flatten().map(|pair| (pair.as_rule(),
                                pair.as_span().start(), pair.as_span().end(), pair.as_str().to_owned()))
                                .collect::<Vec<_>>()).map_err(|error| format!("{error:?}")))));
                    assert_eq!(run(true), run(false), "{rule:?} {source:?} {options:?}");
                }
            }
        }
    }

    #[test]
    fn top_level_units_preserve_absolute_captures_and_merge_pratt_receipts() {
        let source = "/*lead*/;\n#define K 3\ntypedef int T;\nint f(T x){return x+2*3;}\n;int g(int y){return y-4/2;} /*tail*/";
        with_direct_ast(false, || with_native_pratt(true, || with_backend(Backend::Native, || {
            let expected = grameo::engine::parse_program(&native_grammar::Parser,
                Rule::program.name(), source, Default::default()).unwrap()
                .into_typed_pairs::<Rule>().next().unwrap().into_inner()
                .filter(|pair| pair.as_rule() != Rule::EOI).collect::<Vec<_>>();
            let (actual, counts) = profile_construction(|| parse_top_level_units(source, Default::default()));
            let actual = actual.unwrap();
            let flatten = |items: &[grameo::compat::Pair<'_, Rule>]| {
                items.iter().cloned().flat_map(|pair| std::iter::once(pair.clone())
                    .chain(pair.into_inner().flatten())).map(|pair|
                    (pair.as_rule(), pair.as_span().start(), pair.as_span().end(), pair.as_str().to_owned()))
                    .collect::<Vec<_>>()
            };
            assert_eq!(flatten(&actual), flatten(&expected));
            assert_eq!(counts.source_pratt_recaptures, 0, "{counts:?}");
            for item in actual {
                for pair in std::iter::once(item.clone()).chain(item.into_inner().flatten()) {
                    if pair.as_rule() == Rule::binary_expression && pair.as_str().contains(['+', '-']) {
                        let (value, counts) = profile_construction(|| construct_contextual_binary(&pair,
                            |leaf| vybe_ast::Expression::new(vybe_ast::ExprKind::Ident(leaf.as_str().into()))));
                        assert!(value.is_some(), "missing merged receipt at {:?}", pair.as_span());
                        assert_eq!(counts.source_pratt_handoffs, 1);
                    }
                }
            }
        })));
    }

    #[test]
    fn committed_units_are_consumed_by_the_typed_walker_with_both_handoffs() {
        let source = "static int first[]={1,2}; int f(int x){return first[x]+x*2;} \
            static int second[]={3,4}; int g(int x){return second[x]+f(x);} \
            int main(void){return f(1)+g(0);}";
        let reference = crate::parse_with_peg(source).unwrap();
        let ((actual, construction), matching) = with_native_pratt(true, || profile_matching(||
            profile_construction(|| crate::parse_optimized(source))));
        assert_eq!(format!("{:?}", actual.unwrap()), format!("{reference:?}"));
        assert!(matching.rules.iter().any(|row| row.rule == Rule::frontend_unit.name()
            && row.totals.calls > 1), "walker did not consume committed units: {matching:?}");
        assert!(!matching.rules.iter().any(|row| row.rule == Rule::program.name()),
            "valid committed units recaptured the whole program: {matching:?}");
        assert!(construction.source_pratt_handoffs > 0, "{construction:?}");
        assert!(construction.direct_initializer_handoffs > 0, "{construction:?}");
        assert_eq!(construction.declaration_handoffs, 0, "{construction:?}");
    }

    #[test]
    fn committed_units_preserve_standard_wasm_execution_and_noeq() {
        crate::register();
        with_native_pratt(true, || unit_compatibility::execute_identical_standard_wasm(
            crate::profile_source(), [crate::parse_with_peg, crate::parse_optimized]));
    }

    #[test]
    #[ignore = "complete saved C module emission/validation/noEq; provide VYBE_C_FRONTEND_PREPROCESSED"]
    fn committed_units_emit_complete_valid_standard_wasm_without_dynamic_equality() {
        crate::register();
        let path = std::env::var("VYBE_C_FRONTEND_PREPROCESSED")
            .expect("provide the complete saved preprocessed C source");
        let source = std::fs::read_to_string(&path).unwrap();
        let started = std::time::Instant::now();
        let ((module, construction), matching) = with_native_pratt(true, ||
            profile_matching(|| profile_construction(|| crate::parse_optimized(&source))));
        let module = module.unwrap();
        let parsed = started.elapsed();
        eprintln!("complete_standard_wasm_phase parsed elapsed={parsed:?}");
        eprintln!("complete_standard_wasm_matching calls={} submitted_bytes={} failures={} inclusive_elapsed={:?}",
            matching.total.calls, matching.total.bytes, matching.total.failures, matching.total.elapsed);
        assert!(construction.source_pratt_matches > 1, "{construction:?}");
        assert!(matching.rules.iter().any(|row| row.rule == Rule::frontend_unit.name()
            && row.totals.calls > 1), "{matching:?}");
        assert!(!matching.rules.iter().any(|row| row.rule == Rule::program.name()),
            "complete module fell back to whole-program matching: {matching:?}");
        let statements = module.body.len();
        let profile = vybe_runtime::profile::parse_profile_shared(crate::profile_source()).unwrap();
        let mut emission = vybe_compiler::primitives::Compiler::with_shared_profile(profile)
            .begin_emission_owned(module).unwrap();
        let mut emitted_units = 0;
        while emission.emit_next().unwrap().is_some() { emitted_units += 1; }
        assert!(emitted_units > 0);
        let chunks = emission.finish().unwrap().chunks;
        let compiled = started.elapsed();
        let bytecode_bytes: usize = chunks.iter().map(|chunk| chunk.code.len()).sum();
        let source_direct_calls: usize = chunks.iter().map(|chunk|
            chunk.direct_defined_calls.values().filter(|edge| edge.source).count()).sum();
        eprintln!("complete_standard_wasm_phase emitted elapsed={compiled:?} chunks={} bytecode_bytes={bytecode_bytes} source_direct_calls={source_direct_calls}", chunks.len());
        let bytes = vybe_platform_wasm::write_wasm_with_options(&chunks,
            vybe_platform_wasm::WasmWriteOptions {
                include_vybe_metadata: false, include_name_section: true,
            });
        let written = started.elapsed();
        eprintln!("complete_standard_wasm_phase serialized elapsed={written:?} bytes={}", bytes.len());
        wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all())
            .validate_all(&bytes).unwrap();
        unit_compatibility::assert_no_dynamic_equality(&chunks, &bytes);
        if let Some(output) = std::env::var_os("VYBE_C_FRONTEND_WASM_ARTIFACT") {
            std::fs::write(output, &bytes).unwrap();
        }
        eprintln!("complete_standard_wasm input={path} source_bytes={} statements={statements} emitted_units={emitted_units} chunks={} wasm_bytes={} parse={parsed:?} emission={:?} writer={:?} validation_and_audit={:?}",
            source.len(), chunks.len(), bytes.len(), compiled - parsed,
            written - compiled, started.elapsed() - written);
        eprintln!("complete_standard_wasm_construction {construction:?}");
    }

    #[test]
    fn user_call_bindings_outrank_adapters_but_extern_prototypes_do_not() {
        for (source, expected_calls) in [
            ("int read(void){return 7;} int main(void){return read();}", 1),
            ("int read(void); int main(void){return read();} int read(void){return 7;}", 1),
            ("extern long read(int,void*,unsigned long); int main(void){return read(0,0,0);}", 0),
        ] {
            for parse in [crate::parse_with_peg, crate::parse_optimized] {
                let module = with_native_pratt(true, || parse(source)).unwrap();
                let mut calls = 0;
                for statement in &module.body {
                    statement.walk_exprs(&mut |value| {
                        if let vybe_ast::ExprKind::Call { callee, .. } = &value.kind
                            && matches!(&callee.kind, vybe_ast::ExprKind::Ident(name) if name == "read")
                        {
                            calls += 1;
                        }
                    });
                }
                assert_eq!(calls, expected_calls, "{source}");
            }
        }
    }

    #[test]
    fn user_call_bindings_preserve_standard_wasm_execution_and_noeq() {
        crate::register();
        with_native_pratt(true, || unit_compatibility::execute_standard_wasm_programs(
            crate::profile_source(), [crate::parse_with_peg, crate::parse_optimized],
            unit_compatibility::USER_CALL_BINDING_PROGRAMS.iter().copied()));
    }

    #[test]
    fn top_level_units_merge_direct_ownership_and_clear_late_failures() {
        with_native_pratt(true, || with_backend(Backend::Native, || {
            let source = "int a[]={1,2}; int f(int x){return x+2;} int b[]={3,4}; int g(int x){return x*3;}";
            let units = parse_top_level_units(source, Default::default()).unwrap();
            let captures = units.into_iter().flat_map(|pair| std::iter::once(pair.clone())
                .chain(pair.into_inner().flatten())).collect::<Vec<_>>();
            let lists = captures.iter().filter(|pair| pair.as_rule() == Rule::initializer_list).collect::<Vec<_>>();
            assert_eq!(lists.len(), 2);
            for list in lists {
                assert_eq!(list.clone().into_inner().count(), 0);
                assert!(take_initializer_list(list).is_some());
            }
            for pair in captures.iter().filter(|pair| pair.as_rule() == Rule::binary_expression
                && pair.as_str().contains(['+', '*'])) {
                let (_, counts) = profile_construction(|| construct_contextual_binary(pair,
                    |leaf| vybe_ast::Expression::new(vybe_ast::ExprKind::Ident(leaf.as_str().into()))));
                assert_eq!(counts.source_pratt_handoffs, 1);
            }
            let source = "int a[]={1,2}; int f(int x){return x+2;} int broken( {";
            let first = grameo::engine::parse_program_at(&native_grammar::Parser,
                Rule::frontend_unit.name(), source, 0, Default::default()).unwrap()
                .into_typed_pairs::<Rule>().flatten()
                .find(|pair| pair.as_rule() == Rule::initializer_list).unwrap();
            assert!(parse_top_level_units(source, Default::default()).is_err());
            assert!(take_initializer_list(&first).is_none(), "late failure retained an earlier Direct AST");
            let key = (source.as_ptr() as usize, source.len());
            PRATT_CACHE.with(|slot| assert!(!slot.borrow().as_ref().unwrap().contains_key(&key)));
            let original = grameo::engine::parse_program(&native_grammar::Parser,
                Rule::program.name(), source, Default::default()).unwrap_err();
            assert_eq!(format!("{original:?}"), format!("{:?}",
                parse_top_level_units(source, Default::default()).err().expect("late syntax failure")));
        }));
    }

    #[test]
    fn top_level_units_keep_whole_program_errors_and_finite_budgets() {
        let shape = |items: Vec<grameo::compat::Pair<'_, Rule>>| items.into_iter()
            .flat_map(|pair| std::iter::once(pair.clone()).chain(pair.into_inner().flatten()))
            .map(|pair| (pair.as_rule(), pair.as_span().start(), pair.as_span().end(), pair.as_str().to_owned()))
            .collect::<Vec<_>>();
        for source in ["", " /* only trivia */ \n", "; ; int x; ;",
            "int a[]={1,2}; int f(int x){return x+2;}",
            "int a[]={1,2}; int f(int x){return x+;} int broken( {"] {
            for max_steps in (0..128).chain([512, 4096, usize::MAX]) {
                for max_rule_depth in [0, 1, grameo::MatchOptions::default().max_rule_depth] {
                    let options = grameo::MatchOptions { max_steps, max_rule_depth };
                    with_direct_ast(false, || with_native_pratt(true, || with_backend(Backend::Native, || {
                        let expected = parse_captures(Rule::program, source, options)
                            .map(|mut pairs| shape(pairs.next().unwrap().into_inner()
                                .filter(|pair| pair.as_rule() != Rule::EOI).collect()))
                            .map_err(|error| format!("{error:?}"));
                        let actual = parse_top_level_units(source, options).map(&shape)
                            .map_err(|error| format!("{error:?}"));
                        assert_eq!(actual, expected, "{source:?} {options:?}");
                    })));
                }
            }
        }
    }

    #[test]
    fn source_pratt_receipts_keep_complete_typed_modules_and_rejections() {
        assert!(owned::source_pratt_capture_contract(&optimized_grammar::Parser),
            "regenerate the optimized C parser after updating its operator anchors");
        let sources = [
            "typedef short T; int f(T *p, int i){return p[i++]+sizeof(T)*3-2;}",
            "int f(int a,int b,int c,int d){return a+b*c-d/(a+1);}",
            "typedef unsigned long U; U f(U a,U b){return (a<<b)+(a/b)-(a%b);}",
            "int side(int); int f(int a,int b){return side(a++)+side(++b)*side(a+b);}",
            "int f(int *p,int *q){return (p-q)+((p<q)&&(p!=0));}",
            "struct S{int x;}; int f(struct S *p){return p->x+(p->x=3)*2;}",
            "int f(int (*cb)(int),int x){return cb(x)+cb(x+1)*cb(x+2);}",
            "int f(int n,int a[n]){return sizeof(int[n])+a[n-1]*2;}",
            "typedef int T; int f(int x){return sizeof(T)+({int T=3; T+x;})*x;}",
            "double f(double a,double b){return a/b+b*2.0-1.0;}",
            "int f(int a,int b){return a /*l*/ + /*r*/ b*3;}",
            "int f(int a,int b){return (a?b:a+1)+(a,b)*2;}",
            "int f(int a){int x=1+2*3; return x+a*2;}",
            "int a[]={1,2,3}; int f(int x){return a[x]+x*2;}",
            "int f(int a){return a+;}",
            "int f(int a){return a*(2+);}",
            "int f(int a){return a+2;} int broken( {",
        ];
        for source in sources {
            let reference = crate::parse_with_peg(source).map(|value| format!("{value:?}"));
            let old = with_source_pratt(false, || crate::parse_contextual_pratt(source))
                .map(|value| format!("{value:?}"));
            let (actual, profile) = profile_construction(|| with_direct_ast(false,
                || with_source_pratt(true, || crate::parse_contextual_pratt(source))));
            assert_eq!(old, reference, "old Pratt: {source}");
            assert_eq!(actual.map(|value| format!("{value:?}")), reference, "source Pratt: {source}");
            assert!(profile.source_pratt_matches > 0, "{profile:?}: {source}");
            if reference.is_ok() {
                assert!(profile.source_pratt_regions > 0, "{profile:?}: {source}");
                assert!(profile.source_pratt_handoffs > 0, "{profile:?}: {source}");
                // The walker may issue intentionally unsuccessful fragment
                // probes even for a valid module. Its whole-program match
                // must succeed directly, without diagnostic recapture.
                let (pairs, matching) = profile_construction(|| with_backend(Backend::ContextualPratt,
                    || parse_captures(Rule::program, source, Default::default())));
                assert!(pairs.is_ok());
                assert_eq!(matching.source_pratt_matches, 1);
                assert_eq!(matching.source_pratt_recaptures, 0, "{matching:?}: {source}");
            }
        }
    }

    #[test]
    fn source_pratt_receipts_match_complete_ast_for_all_operator_pairs() {
        let operators = ["||", "&&", "|", "^", "&", "==", "!=", "<", "<=", ">", ">=",
            "<<", ">>", "+", "-", "*", "/", "%"];
        for left in operators {
            for right in operators {
                let source = format!("int f(int a,int b,int c){{return a {left} b {right} c;}}");
                let reference = crate::parse_with_peg(&source).unwrap();
                let (actual, profile) = profile_construction(|| with_direct_ast(false,
                    || with_source_pratt(true, || crate::parse_contextual_pratt(&source))));
                assert_eq!(format!("{:?}", actual.unwrap()), format!("{reference:?}"), "{source}");
                assert!(profile.source_pratt_handoffs > 0, "{profile:?}: {source}");
            }
        }
    }

    #[test]
    fn source_pratt_receipts_rebind_and_survive_reentrant_cache_invalidation() {
      for backend in [Backend::ContextualPratt, Backend::Native] {
        with_native_pratt(true, || with_backend(backend, || {
            let source = "a+b*c+d";
            let pair = parse_captures(Rule::binary_expression, source, Default::default())
                .unwrap().next().unwrap();
            for binding in [1, 7, 1] {
                let mut order = Vec::new();
                let (value, counts) = profile_construction(|| construct_contextual_binary(&pair, |child| {
                    order.push(child.as_str().to_owned());
                    // Replacing the TLS map must not invalidate this live recipe.
                    if order.len() == 1 {
                        assert!(parse_captures(Rule::binary_expression, source,
                            grameo::MatchOptions { max_steps: 1, ..Default::default() }).is_err());
                    }
                    vybe_ast::Expression::new(vybe_ast::ExprKind::Lit(vybe_ast::Literal::Int(binding)))
                }));
                let expected = owned::construct_contextual_binary(&pair, |_| {
                    vybe_ast::Expression::new(vybe_ast::ExprKind::Lit(vybe_ast::Literal::Int(binding)))
                }).unwrap();
                assert_eq!(format!("{:?}", value.unwrap()), format!("{expected:?}"));
                assert_eq!(order, ["a", "b", "c", "d"]);
                assert_eq!(counts.source_pratt_handoffs, 1);
                // Recreate receipts for the next independent scope binding.
                parse_captures(Rule::binary_expression, source, Default::default()).unwrap();
            }
        }));
      }
    }

    #[test]
    fn source_pratt_and_direct_initializers_share_one_transactional_match() {
      for backend in [Backend::ContextualPratt, Backend::Native] {
        with_native_pratt(true, || with_backend(backend, || {
            let source = "int a[]={1,2,3}; int f(int x){return a[x]+x*2;}";
            let (pairs, counts) = profile_construction(||
                parse_captures(Rule::program, source, Default::default()));
            let captures: Vec<_> = pairs.unwrap().flatten().collect();
            assert_eq!(counts.parse_calls, 1);
            assert_eq!(counts.source_pratt_matches, 1);
            assert_eq!(counts.source_pratt_recaptures, 0);
            if backend == Backend::Native { assert_eq!(counts.declaration_regions, 0); }
            let list = captures.iter().find(|pair| pair.as_rule() == Rule::initializer_list).unwrap();
            assert_eq!(list.clone().into_inner().count(), 0,
                "a certified Direct initializer must not retain descendants for inactive Pratt frames");
            let value = take_initializer_list(list).expect("same-match Direct initializer publication");
            assert!(matches!(&value.kind, vybe_ast::ExprKind::Array(values) if values.len() == 3));
            let region = captures.iter().find(|pair| pair.as_rule() == Rule::binary_expression
                && pair.as_str().contains("a[x]")).unwrap();
            let (value, counts) = profile_construction(|| construct_contextual_binary(region,
                |child| vybe_ast::Expression::new(vybe_ast::ExprKind::Ident(child.as_str().into()))));
            assert!(value.is_some());
            assert_eq!(counts.source_pratt_handoffs, 1);
            assert!(parse_captures(Rule::program, source,
                grameo::MatchOptions { max_steps: 1, ..Default::default() }).is_err());
            assert!(take_initializer_list(list).is_none(), "failed composed reparse must invalidate Direct cache");
        }));
      }
    }

    #[test]
    #[cfg(unix)]
    #[ignore = "same-cohort full frontend AST parity and rotated source-Pratt receipt timing"]
    fn measure_source_pratt_full_frontend() {
        measure_pratt_full_frontend(false);
    }

    #[test]
    #[cfg(unix)]
    #[ignore = "same-cohort Native-only full frontend parity and declaration-free Pratt timing"]
    fn measure_native_pratt_full_frontend() {
        measure_pratt_full_frontend(true);
    }

    #[cfg(unix)]
    fn measure_pratt_full_frontend(native_only: bool) {
        assert!(std::env::var_os("VYBE_C_DEBUG_TIMINGS").is_none());
        assert!(owned::source_pratt_capture_contract(&optimized_grammar::Parser));
        let path = std::env::var("VYBE_C_FRONTEND_PREPROCESSED").unwrap();
        let source = std::fs::read_to_string(&path).unwrap();
        let reference = format!("{:?}", crate::parse_with_peg(&source).unwrap());
        let run = |index| if native_only {
            with_direct_ast(index % 2 == 1, || with_native_pratt(index >= 2,
                || crate::parse_optimized(&source)))
        } else { match index {
            0 => with_native_pratt(false, || with_direct_ast(false, || crate::parse_optimized(&source))),
            1 => with_native_pratt(false, || with_direct_ast(true, || crate::parse_optimized(&source))),
            2 | 3 => with_direct_ast(true, || with_source_pratt(index == 3,
                || crate::parse_contextual_pratt(&source))),
            _ => unreachable!(),
        }};
        let labels = if native_only {
            ["NativeCapture", "NativeDirect", "NativePrattCapture", "NativePrattDirect"]
        } else { ["NativeCapture", "NativeDirect", "CapturePratt+Direct", "SourcePratt+Direct"] };
        for index in 0..4 {
            let (actual, counts) = profile_construction(|| run(index));
            let actual = actual.unwrap();
            let actual_text = format!("{actual:?}");
            assert!(actual_text == reference, "{} differs from full PEG AST at byte {:?}; expected={} actual={}",
                labels[index], actual_text.bytes().zip(reference.bytes()).position(|(a,b)| a != b),
                reference.len(), actual_text.len());
            if index == 3 || native_only && index == 2 { assert!(counts.source_pratt_handoffs > 0, "{counts:?}"); }
            else { assert_eq!(counts.source_pratt_handoffs, 0); }
            if native_only { assert_eq!(counts.declaration_handoffs, 0); }
            eprintln!("source_pratt_parity input={path} bytes={} statements={} backend={} {counts:?}",
                source.len(), actual.body.len(), labels[index]);
        }
        drop(reference);
        let samples = std::env::var("VYBE_C_FRONTEND_SAMPLES").ok()
            .and_then(|value| value.parse::<usize>().ok()).unwrap_or(6);
        assert!(samples == 0 || samples >= 3, "zero selects parity-only");
        if samples == 0 { return; }
        let mut cpus: [Vec<u128>; 4] = std::array::from_fn(|_| Vec::new());
        let mut walls: [Vec<u128>; 4] = std::array::from_fn(|_| Vec::new());
        let orders = [[0, 1, 2, 3], [1, 2, 3, 0], [2, 3, 0, 1], [3, 0, 1, 2],
            [3, 2, 1, 0], [0, 3, 2, 1], [1, 0, 3, 2], [2, 1, 0, 3]];
        for sample in 0..samples {
            for index in orders[sample % orders.len()] {
                let wall = Instant::now();
                let cpu = cpu_ns();
                let module = run(index).unwrap();
                drop(std::hint::black_box(module));
                let cpu = cpu_ns() - cpu;
                let wall = wall.elapsed().as_nanos();
                cpus[index].push(cpu);
                walls[index].push(wall);
                eprintln!("source_pratt_sample input={path} sample={sample} backend={} cpu_ns={cpu} wall_ns={wall}", labels[index]);
            }
            eprintln!("source_pratt_pair sample={sample} source_over_capture_pratt_cpu={} source_over_native_direct_cpu={} source_over_native_capture_cpu={} direct_over_native_capture_cpu={}",
                cpus[3][sample] as f64 / cpus[2][sample] as f64,
                cpus[3][sample] as f64 / cpus[1][sample] as f64,
                cpus[3][sample] as f64 / cpus[0][sample] as f64,
                cpus[1][sample] as f64 / cpus[0][sample] as f64);
        }
        for index in 0..4 {
            cpus[index].sort_unstable(); walls[index].sort_unstable();
            eprintln!("source_pratt_median input={path} backend={} samples={samples} cpu_ns={} wall_ns={}",
                labels[index], cpus[index][samples / 2], walls[index][samples / 2]);
        }
    }

    #[test]
    fn identifier_probes_keep_semantic_builder_captures_errors_and_limits() {
        let flatten = |result: Result<grameo::ParseTree<'_, '_>, grameo::ParseError>| {
            result.map(|tree| tree.into_typed_pairs::<Rule>().flatten()
                .map(|pair| (pair.as_rule(), pair.as_span().start(), pair.as_span().end(),
                    pair.as_str().to_owned()))
                .collect::<Vec<_>>()).map_err(|error| format!("{error:?}"))
        };
        for source in ["x", "_private42", "int", "sizeof", "x+y", "42", "", " x", "é", "x/*tail*/"] {
            for max_steps in (0..96).chain(std::iter::once(usize::MAX)) {
                for max_rule_depth in [0, 1, grameo::MatchOptions::default().max_rule_depth] {
                    let options = grameo::MatchOptions { max_steps, max_rule_depth };
                    let native = flatten(build_hybrid(Rule::ident_name, source, options, &mut NativeBuilder));
                    assert_eq!(native, flatten(build_hybrid(Rule::ident_name, source, options,
                        &mut owned::Owned::new(source))), "owned {source:?} {options:?}");
                    assert_eq!(native, flatten(build_hybrid(Rule::ident_name, source, options,
                        &mut owned::ContextualBuilder::new(source))), "mixed {source:?} {options:?}");
                }
            }
        }
    }

    #[test]
    fn identifier_probes_invalidate_only_same_source_without_publishing_empty_caches() {
        for backend in [Backend::Owned, Backend::OwnedContextual, Backend::ContextualPratt] {
            for max_steps in [1, usize::MAX] {
                with_backend(backend, || {
                    let source = "identifier";
                    let key = (source.as_ptr() as usize, source.len());
                    let other = (0, 0);
                    OWNED_CACHE.with(|slot| {
                        if let Some(cache) = slot.borrow_mut().as_mut() {
                            cache.insert(key, HashMap::new());
                            cache.insert(other, HashMap::new());
                        }
                    });
                    MIXED_CACHE.with(|slot| {
                        if let Some(cache) = slot.borrow_mut().as_mut() {
                            cache.insert(key, Rc::new(HashMap::new()));
                            cache.insert(other, Rc::new(HashMap::new()));
                        }
                    });
                    DECLARATION_CACHE.with(|slot| {
                        if let Some(cache) = slot.borrow_mut().as_mut() {
                            cache.insert(key, HashMap::new());
                            cache.insert(other, HashMap::new());
                        }
                    });
                    let result = parse_captures(Rule::ident_name, source,
                        grameo::MatchOptions { max_steps, ..Default::default() });
                    assert_eq!(result.is_ok(), max_steps == usize::MAX);
                    macro_rules! check_cache {
                        ($slot:ident) => {
                            $slot.with(|slot| {
                                if let Some(cache) = slot.borrow().as_ref() {
                                    assert!(!cache.contains_key(&key));
                                    assert!(cache.contains_key(&other));
                                }
                            });
                        };
                    }
                    check_cache!(OWNED_CACHE);
                    check_cache!(MIXED_CACHE);
                    check_cache!(DECLARATION_CACHE);
                });
            }
        }
    }

    fn with_pratt_declarations<T>(enabled: bool, action: impl FnOnce() -> T) -> T {
        struct Restore(bool);
        impl Drop for Restore {
            fn drop(&mut self) {
                PRATT_DECLARATIONS_DISABLED.with(|slot| slot.set(self.0));
            }
        }
        let _restore = Restore(PRATT_DECLARATIONS_DISABLED.with(|slot| slot.replace(!enabled)));
        action()
    }

    #[test]
    fn matching_collector_preserves_capture_spans_errors_and_copy_profile() {
        fn copy_contract<T: Copy>() {}
        copy_contract::<ConstructionProfile>();
        fn flatten(pairs: grameo::compat::Pairs<'_, Rule>, out: &mut Vec<(Rule, usize, usize)>) {
            for pair in pairs {
                out.push((pair.as_rule(), pair.as_span().start(), pair.as_span().end()));
                flatten(pair.into_inner(), out);
            }
        }
        let cases = [
            (Rule::program, "typedef int T; int f(T x){return x+1;}", usize::MAX),
            (Rule::binary_expression, "1+2*3", usize::MAX),
            (Rule::binary_expression, "+", usize::MAX),
            (Rule::program, "int f(void){return 1;}", 1),
        ];
        for backend in [Backend::Peg, Backend::Native, Backend::Owned,
            Backend::OwnedContextual, Backend::ContextualPratt] {
            let run = || with_backend(backend, || cases.iter().map(|&(rule, source, max_steps)| {
                parse_captures(rule, source, grameo::MatchOptions { max_steps, ..Default::default() })
                    .map(|pairs| { let mut out = Vec::new(); flatten(pairs, &mut out); out })
                    .map_err(|error| format!("{error:?}"))
            }).collect::<Vec<_>>());
            let expected = run();
            let (actual, profile) = profile_matching(run);
            assert_eq!(actual, expected, "{backend:?}");
            assert_eq!(profile.total.calls, cases.len());
            assert_eq!(profile.total.bytes, cases.iter().map(|(_, source, _)| source.len()).sum::<usize>());
            assert_eq!(profile.total.failures, 2);
            assert_eq!(profile.rules.len(), 2);
            assert!(profile.rules.iter().all(|row| row.backend == backend));
            assert_eq!(profile.rules.iter().map(|row| row.totals.calls).sum::<usize>(), cases.len());
            assert_eq!(profile.overflow.calls, 0);
            assert!(!MATCHING_ENABLED.with(Cell::get));
            assert!(MATCHING_PROFILE.with(|slot| slot.borrow().is_none()));
        }
    }

    #[test]
    fn matching_collector_is_bounded_and_restores_nested_and_panicked_scopes() {
        let (_, outer) = profile_matching(|| {
            record_matching(Backend::Peg, "outer", 3, Duration::from_nanos(1), false);
            let (_, inner) = profile_matching(|| {
                for index in 0..MATCHING_RULE_LIMIT + 3 {
                    record_matching(Backend::Peg, &format!("rule_{index}"), 2, Duration::from_nanos(1), false);
                }
            });
            assert_eq!(inner.rules.len(), MATCHING_RULE_LIMIT);
            assert_eq!(inner.overflow.calls, 3);
            assert_eq!(inner.overflow.bytes, 6);
            assert_eq!(inner.total.calls, MATCHING_RULE_LIMIT + 3);
            let unwind = std::panic::catch_unwind(|| profile_matching(|| panic!("collector rollback")));
            assert!(unwind.is_err());
            assert!(MATCHING_ENABLED.with(Cell::get));
            record_matching(Backend::Peg, "outer", 5, Duration::from_nanos(2), true);
        });
        assert_eq!(outer.rules.len(), 1);
        assert_eq!(outer.total.calls, 2);
        assert_eq!(outer.total.bytes, 8);
        assert_eq!(outer.total.failures, 1);
        assert_eq!(outer.total.elapsed, Duration::from_nanos(3));
        assert!(!MATCHING_ENABLED.with(Cell::get));
        assert!(MATCHING_PROFILE.with(|slot| slot.borrow().is_none()));
    }

    #[test]
    #[ignore = "single instrumented frontend run with bounded matching attribution"]
    fn profile_fragment_matching_frontend() {
        let path = std::env::var("VYBE_C_FRONTEND_PREPROCESSED")
            .expect("provide a saved preprocessed C source");
        let source = std::fs::read_to_string(&path).unwrap();
        let started = Instant::now();
        let ((result, construction), matching) = profile_matching(|| {
            profile_construction(|| crate::parse(&source))
        });
        let module = result.unwrap();
        eprintln!("fragment_matching_frontend input={path} bytes={} statements={} frontend_wall_ns={} {construction:?}",
            source.len(), module.body.len(), started.elapsed().as_nanos());
        assert_eq!(construction.parse_calls, matching.total.calls);
        for row in &matching.rules {
            eprintln!("fragment_matching_rule backend={:?} rule={} calls={} submitted_bytes={} failures={} matching_wall_ns={}",
                row.backend, row.rule, row.totals.calls, row.totals.bytes, row.totals.failures,
                row.totals.elapsed.as_nanos());
        }
        eprintln!("fragment_matching_total calls={} submitted_bytes={} failures={} inclusive_matching_wall_ns={} overflow={:?}",
            matching.total.calls, matching.total.bytes, matching.total.failures,
            matching.total.elapsed.as_nanos(), matching.overflow);
    }

    #[cfg(unix)]
    fn cpu_ns() -> u128 {
        let mut time = libc::timespec {
            tv_sec: 0,
            tv_nsec: 0,
        };
        assert_eq!(
            unsafe { libc::clock_gettime(libc::CLOCK_THREAD_CPUTIME_ID, &mut time) },
            0
        );
        time.tv_sec as u128 * 1_000_000_000 + time.tv_nsec as u128
    }

    #[test]
    fn backend_restores_nested_calls_and_unwinds() {
        with_backend(Backend::Native, || {
            assert_eq!(BACKEND.with(Cell::get), Some(Backend::Native));
            with_backend(Backend::Peg, || {
                assert_eq!(BACKEND.with(Cell::get), Some(Backend::Peg))
            });
            assert_eq!(BACKEND.with(Cell::get), Some(Backend::Native));
        });
        assert_eq!(BACKEND.with(Cell::get), None);
        let _ = std::panic::catch_unwind(|| with_backend(Backend::Native, || panic!("rollback")));
        assert_eq!(BACKEND.with(Cell::get), None);
    }

    #[test]
    fn committed_mixed_schedules_match_all_operator_pairs() {
        let operators = [
            "||", "&&", "|", "^", "&", "==", "!=", "<", "<=", ">", ">=",
            "<<", ">>", "+", "-", "*", "/", "%",
        ];
        for left in operators {
            for right in operators {
                let source = format!(
                    "int f(int a,int b,int c){{return a {left} b {right} c;}}"
                );
                let expected = crate::parse_with_peg(&source).unwrap();
                let (actual, profile) = profile_construction(|| {
                    crate::parse_owned_contextual(&source).unwrap()
                });
                assert_eq!(format!("{expected:?}"), format!("{actual:?}"), "{source}");
                assert!(profile.mixed_handoffs > 0, "{source}");
                assert!(profile.contextual_operands >= 3, "{source}");
                let (direct, profile) = profile_construction(|| {
                    with_direct_ast(false, || crate::parse_contextual_pratt(&source)).unwrap()
                });
                assert_eq!(format!("{expected:?}"), format!("{direct:?}"), "pratt: {source}");
                assert!(profile.contextual_pratt_handoffs > 0, "{source}");
                assert!(profile.contextual_pratt_operands >= 3, "{source}");
            }
        }
    }

    #[test]
    fn owned_handoff_moves_replays_and_restores_nested_cache() {
        let source = "1+2*3";
        with_backend(Backend::Owned, || {
            let mut pairs = parse_owned_captures(
                Rule::binary_expression,
                source,
                grameo::MatchOptions::default(),
            )
            .unwrap();
            let pair = pairs.next().unwrap();
            let first = take_binary(&pair).unwrap();
            with_backend(Backend::Peg, || assert!(take_binary(&pair).is_none()));
            assert_eq!(
                format!("{first:?}"),
                format!("{:?}", take_binary(&pair).unwrap())
            );
            assert_eq!(
                format!("{first:?}"),
                format!("{:?}", take_binary(&pair).unwrap())
            );
        });
        assert!(OWNED_CACHE.with(|cache| cache.borrow().is_none()));
    }

    #[test]
    fn mixed_regions_rebind_each_operand_in_order_without_cached_scope() {
        let source = "x+1*2==3";
        with_backend(Backend::OwnedContextual, || {
            let pair = parse_captures(Rule::binary_expression, source, Default::default())
                .unwrap()
                .next()
                .unwrap();
            for binding in [4, 7, 4] {
                let mut order = Vec::new();
                let actual = take_binary_with_context(&pair, |operand| {
                    order.push(operand.as_str().to_owned());
                    vybe_ast::Expression::new(vybe_ast::ExprKind::Lit(vybe_ast::Literal::Int(
                        if operand.as_str() == "x" {
                            binding
                        } else {
                            operand.as_str().parse().unwrap()
                        },
                    )))
                })
                .unwrap();
                assert_eq!(order, ["x", "1", "2", "3"]);
                let leaf = |n| {
                    vybe_ast::Expression::new(vybe_ast::ExprKind::Lit(vybe_ast::Literal::Int(n)))
                };
                let binary = |op, left, right| {
                    vybe_ast::Expression::new(vybe_ast::ExprKind::Binary {
                        op,
                        left: Box::new(left),
                        right: Box::new(right),
                    })
                };
                let expected = binary(
                    vybe_ast::BinOp::Eq,
                    binary(
                        vybe_ast::BinOp::Add,
                        leaf(binding),
                        binary(vybe_ast::BinOp::Mul, leaf(1), leaf(2)),
                    ),
                    leaf(3),
                );
                assert_eq!(format!("{actual:?}"), format!("{expected:?}"));
            }
        });
    }

    #[test]
    fn mixed_regions_keep_rich_operands_and_recursive_cache_reads() {
        let source = "p->x+(a?b:c)*f(2)";
        with_backend(Backend::OwnedContextual, || {
            let pair = parse_captures(Rule::binary_expression, source, Default::default())
                .unwrap()
                .next()
                .unwrap();
            let mut order = Vec::new();
            let value = take_binary_with_context(&pair, |operand| {
                // Consumers may reenter matching and scoped cache selection.
                with_backend(Backend::Peg, || {
                    parse_captures(Rule::unary_expression, operand.as_str(), Default::default())
                        .unwrap();
                });
                order.push(operand.as_str().to_owned());
                vybe_ast::Expression::new(vybe_ast::ExprKind::Ident(operand.as_str().to_owned()))
            })
            .unwrap();
            assert_eq!(order, ["p->x", "(a?b:c)", "f(2)"]);
            assert!(matches!(
                value.kind,
                vybe_ast::ExprKind::Binary {
                    op: vybe_ast::BinOp::Add,
                    ..
                }
            ));
            assert!(
                take_binary_with_context(&pair, |operand| vybe_ast::Expression::new(
                    vybe_ast::ExprKind::Ident(operand.as_str().to_owned())
                ))
                .is_some()
            );
        });
    }

    #[test]
    fn mixed_cache_allows_reentrant_regions_and_clears_failed_same_source() {
        let source = "int f(void){return a+b;} int g(void){return c*d;}";
        with_backend(Backend::OwnedContextual, || {
            let pairs = parse_captures(Rule::program, source, Default::default()).unwrap();
            let regions: Vec<_> = pairs
                .flatten()
                .filter(|pair| pair.as_rule() == Rule::binary_expression)
                .collect();
            assert_eq!(regions.len(), 2);
            let mut nested = 0;
            let leaf = |pair: grameo::compat::Pair<'_, Rule>| {
                vybe_ast::Expression::new(vybe_ast::ExprKind::Ident(pair.as_str().to_owned()))
            };
            assert!(
                take_binary_with_context(&regions[0], |pair| {
                    assert!(take_binary_with_context(&regions[1], leaf).is_some());
                    nested += 1;
                    leaf(pair)
                })
                .is_some()
            );
            assert_eq!(nested, 2);
            assert!(
                parse_captures(
                    Rule::program,
                    source,
                    grameo::MatchOptions {
                        max_steps: 1,
                        ..Default::default()
                    }
                )
                .is_err()
            );
            assert!(take_binary_with_context(&regions[0], leaf).is_none());
            assert!(take_binary_with_context(&regions[1], leaf).is_none());
        });
    }

    #[test]
    fn mixed_nested_operators_belong_to_their_own_capture_region() {
        let source = "f(a+b)*(c-d)+x[2*3]";
        with_backend(Backend::OwnedContextual, || {
            let reference = with_backend(Backend::Peg, || {
                parse_captures(Rule::binary_expression, source, Default::default()).unwrap()
            });
            let candidate =
                parse_captures(Rule::binary_expression, source, Default::default()).unwrap();
            let captures = |pairs: grameo::compat::Pairs<'_, Rule>| {
                pairs
                    .flatten()
                    .map(|pair| (pair.as_rule(), pair.as_span().start(), pair.as_span().end()))
                    .collect::<Vec<_>>()
            };
            assert_eq!(captures(candidate.clone()), captures(reference));
            let regions: Vec<_> = candidate
                .flatten()
                .filter(|pair| pair.as_rule() == Rule::binary_expression)
                .collect();
            for region in regions {
                let expected: Vec<_> = region
                    .clone()
                    .into_inner()
                    .filter(|pair| pair.as_rule() == Rule::unary_expression)
                    .map(|pair| pair.as_str().to_owned())
                    .collect();
                let mut actual = Vec::new();
                let result = take_binary_with_context(&region, |pair| {
                    actual.push(pair.as_str().to_owned());
                    vybe_ast::Expression::new(vybe_ast::ExprKind::Ident(pair.as_str().into()))
                });
                if expected.len() > 1 {
                    assert!(result.is_some(), "{}", region.as_str());
                    assert_eq!(actual, expected);
                } else {
                    assert!(result.is_none());
                    assert!(actual.is_empty());
                }
            }
        });
    }

    #[test]
    fn mixed_reentrant_failed_match_cannot_resurrect_a_recipe() {
        let source = "x+y";
        with_backend(Backend::OwnedContextual, || {
            let pair = parse_captures(Rule::binary_expression, source, Default::default())
                .unwrap()
                .next()
                .unwrap();
            let leaf = |pair: grameo::compat::Pair<'_, Rule>| {
                vybe_ast::Expression::new(vybe_ast::ExprKind::Ident(pair.as_str().into()))
            };
            let mut invalidated = false;
            let result = take_binary_with_context(&pair, |operand| {
                if !invalidated {
                    assert!(
                        parse_captures(
                            Rule::binary_expression,
                            source,
                            grameo::MatchOptions {
                                max_steps: 1,
                                ..Default::default()
                            }
                        )
                        .is_err()
                    );
                    invalidated = true;
                }
                leaf(operand)
            });
            assert!(result.is_some());
            assert!(invalidated);
            assert!(take_binary_with_context(&pair, leaf).is_none());
        });
    }

    #[test]
    fn declaration_cache_moves_values_and_failed_reparse_clears_them() {
        let source = "int x=2; int y=3;";
        for backend in [Backend::OwnedContextual, Backend::ContextualPratt] {
        with_backend(backend, || {
            let pairs = parse_captures(Rule::program, source, Default::default()).unwrap();
            let declarations: Vec<_> = pairs
                .flatten()
                .filter(|p| p.as_rule() == Rule::init_declarator)
                .collect();
            assert_eq!(declarations.len(), 2);
            assert!(take_scalar_declarator(&declarations[0]).is_some());
            assert!(take_scalar_declarator(&declarations[0]).is_none());
            assert!(
                parse_captures(
                    Rule::program,
                    source,
                    grameo::MatchOptions {
                        max_steps: 1,
                        ..Default::default()
                    }
                )
                .is_err()
            );
            assert!(take_scalar_declarator(&declarations[1]).is_none());
        });
        }
    }

    fn with_plain_declarator_scan<T>(enabled: bool, action: impl FnOnce() -> T) -> T {
        struct Restore(bool);
        impl Drop for Restore {
            fn drop(&mut self) {
                PLAIN_DECLARATOR_SCAN_DISABLED.with(|flag| flag.set(self.0));
            }
        }
        let _restore = Restore(PLAIN_DECLARATOR_SCAN_DISABLED.with(|flag| flag.replace(!enabled)));
        action()
    }

    #[test]
    fn plain_scalar_receipt_skips_scans_independently_per_comma_item() {
        for backend in [Backend::OwnedContextual, Backend::ContextualPratt] {
            for (source, expected_bypasses, expected_fallbacks) in [
                ("int f(void){int a=1,b=2; return a+b;}", 2, 0),
                ("int f(void){int a=1,*p=0,b[2]={1,2},c=3; return a+c;}", 2, 2),
                ("typedef int *P; int f(void){P p=0; int value=2; return value;}", 1, 1),
                ("typedef int A[2]; int f(void){A a; int value=2; return value;}", 1, 1),
                ("int f(void){int a=1,b=a+1,c=3; return b+c;}", 2, 1),
            ] {
                let expected = crate::parse_with_peg(source);
                let (reference, reference_counts) = profile_construction(|| {
                    with_plain_declarator_scan(false, || with_backend(backend,
                        || crate::walker::parse(source)))
                });
                let (actual, counts) = profile_construction(|| {
                    with_backend(backend, || crate::walker::parse(source))
                });
                assert!(expected.is_ok(), "{source}: {expected:?}");
                assert_eq!(format!("{reference:?}"), format!("{expected:?}"), "{backend:?}: {source}");
                assert_eq!(format!("{actual:?}"), format!("{expected:?}"), "{backend:?}: {source}");
                assert_eq!(counts.declarator_scan_bypasses, expected_bypasses, "{backend:?}: {source}");
                assert_eq!(counts.declarator_scan_fallbacks, expected_fallbacks, "{backend:?}: {source}");
                assert_eq!(reference_counts.declarator_scan_bypasses, 0);
                assert_eq!(reference_counts.declarator_scan_fallbacks,
                    counts.declarator_scan_fallbacks + counts.declarator_scan_bypasses);
                assert_eq!(reference_counts.declaration_handoffs, counts.declaration_handoffs);
            }
        }
    }

    #[test]
    fn plain_scalar_receipt_preserves_types_storage_scope_and_errors() {
        for source in [
            "typedef unsigned short T; const T global=65535; int f(void){T a=2,b; {int T=3; b=T;} return a+b;}",
            "typedef int T; int f(void){T /*name*/ a /*init*/ =2, b=3; return a+b;}",
            "int global=2; extern int global; int f(void){static int n=3; volatile int v=4; register int r=5; return global+n+v+r;}",
            "float f(void){float a=1.25f; double b=2.5; return a+b;}",
            "typedef int F(int); typedef int (*FP)(int); F fn; int f(void){FP p=0; return p!=0;}",
            "typedef int T; int f(void){int T=2; T*=3; return T;}",
            "int f(void){unsigned a=4294967295u; unsigned long long b=18446744073709551615ULL; return a!=b;}",
            "int f(void){int a=1,b=; return a;}",
            "int f(void){int a=1,*p=; return a;}",
            "int f(void){int a=1,b=(2+); return a;}",
            "typedef int T; int f(void){T a=1; return (T;}",
        ] {
            let expected = crate::parse_with_peg(source);
            for backend in [Backend::OwnedContextual, Backend::ContextualPratt] {
                let reference = with_plain_declarator_scan(false,
                    || with_backend(backend, || crate::walker::parse(source)));
                let actual = with_backend(backend, || crate::walker::parse(source));
                assert_eq!(format!("{reference:?}"), format!("{expected:?}"), "{backend:?}: {source}");
                assert_eq!(format!("{actual:?}"), format!("{expected:?}"), "{backend:?}: {source}");
            }
        }
    }

    #[test]
    fn plain_scalar_receipt_keeps_source_spans_and_moves_once() {
        let source = "int a /*gap*/ = 2, *p=0, b, c=3;";
        for backend in [Backend::OwnedContextual, Backend::ContextualPratt] {
            with_backend(backend, || {
                let pairs = parse_captures(Rule::program, source, Default::default()).unwrap();
                let mut found = Vec::new();
                for pair in pairs.flatten().filter(|pair| pair.as_rule() == Rule::init_declarator) {
                    if let Some(receipt) = take_scalar_declarator(&pair) {
                        let syntax = receipt.syntax;
                        let span = pair.as_span();
                        assert_eq!(syntax.declarator.start, span.start());
                        assert!(syntax.declarator.end <= span.end());
                        let name = &source[syntax.declarator.start..syntax.declarator.end];
                        assert!(matches!(&receipt.declaration.pattern,
                            vybe_ast::BindingPattern::Ident(actual) if actual == name));
                        let initializer = syntax.initializer.map(|span| &source[span.start..span.end]);
                        assert_eq!(initializer.is_some(), receipt.declaration.init.is_some());
                        found.push((name, initializer));
                        assert!(take_scalar_declarator(&pair).is_none());
                    } else {
                        assert!(pair.as_str().starts_with('*'));
                    }
                }
                assert_eq!(found, [("a", Some("2")), ("b", None), ("c", Some("3"))]);
            });
        }
    }

    #[test]
    fn contextual_pratt_composes_declarations_without_mixed_replay() {
        let mut direct_declarations = 0;
        for source in [
            "int seed=3; int f(int x){int y=2; return x+y*seed;}",
            "typedef unsigned short word; int f(word x){int y=2; {int word=3; y+=word;} return x+y;}",
            "int f(int x){static int count=2; const int step=3; count+=x; return count+step;}",
            "struct P { int x; }; int f(struct P *p){int a[2]={1,2}; int n=3; return p->x+a[1]*n;}",
        ] {
            let expected = crate::parse_with_peg(source).unwrap();
            for direct in [false, true] {
            let baseline = with_direct_ast(direct, || with_pratt_declarations(false,
                || crate::parse_contextual_pratt(source))).unwrap();
            let (actual, counts) = with_direct_ast(direct, ||
                profile_construction(|| crate::parse_contextual_pratt(source)));
            assert_eq!(format!("{:?}", actual.unwrap()), format!("{expected:?}"), "{source}");
            assert_eq!(format!("{baseline:?}"), format!("{expected:?}"), "{source}");
            assert!(counts.declaration_regions > 0, "{source}: {counts:?}");
            assert!(counts.declaration_handoffs + counts.direct_declaration_handoffs > 0,
                "{source}: {counts:?}");
            if direct {
                direct_declarations += counts.direct_declaration_handoffs;
            } else {
                assert!(counts.declaration_handoffs > 0, "{source}: {counts:?}");
                assert_eq!(counts.direct_declaration_handoffs, 0, "{counts:?}");
            }
            assert!(counts.contextual_pratt_handoffs + counts.direct_binary_handoffs > 0,
                "{source}: {counts:?}");
            assert_eq!(counts.mixed_regions, 0);
            assert_eq!(counts.mixed_handoffs, 0);
            }
        }
        assert!(direct_declarations > 0, "combined path did not consume Direct declarations");
    }

    #[test]
    fn staged_declarations_reject_contextual_shapes_without_losing_errors() {
        for source in [
            "int *x=2;",
            "int x[2];",
            "int x=sizeof(int);",
            "int x=1(2);",
            "int x=a+3;",
            "int x=(int)2;",
            "int x={2};",
        ] {
            let mut context = DeclarationContext::new(source);
            context.declarations.insert(
                (0, source.len()),
                DeclarationPolicy {
                    kind: vybe_ast::VarDeclKind::Var,
                    statement_span: Default::default(),
                },
            );
            let start = source.find('x').unwrap();
            // Give the target a policy even for unsupported shapes. Rejection
            // must be structural, not a consequence of a missing policy.
            for start in [start, source[..start].rfind('*').unwrap_or(start)] {
                context.declarators.insert(
                    (start, source.len() - 1),
                    ScalarPolicy {
                        emitted_name: "x".into(),
                        type_hint: Some("int".into()),
                        zero_if_missing: true,
                        integer_wrap: None,
                    },
                );
            }
            assert!(
                construct_scalar_declaration(source, context, Default::default())
                    .unwrap()
                    .is_none(),
                "{source}"
            );
        }
        let source = String::from("int x=2;");
        let other = source.clone();
        assert!(
            construct_scalar_declaration(
                &source,
                DeclarationContext::new(&other),
                Default::default()
            )
            .is_err()
        );
        assert!(
            construct_scalar_declaration(
                "int x=;",
                DeclarationContext::new("int x=;"),
                Default::default()
            )
            .is_err()
        );
    }

    #[test]
    fn staged_scalar_literals_and_prefixes_construct_without_recapture() {
        for (source, raw, rule) in [
            ("double x=1.25;", "1.25", Rule::float_literal),
            ("int x='a';", "'a'", Rule::char_literal),
            ("int x=L'\\n';", "L'\\n'", Rule::wide_char_literal),
            ("int x=true;", "true", Rule::bool_literal),
            ("int x=-2;", "2", Rule::int_literal),
        ] {
            let mut context = DeclarationContext::new(source);
            context.declarations.insert(
                (0, source.len()),
                DeclarationPolicy {
                    kind: vybe_ast::VarDeclKind::Var,
                    statement_span: Default::default(),
                },
            );
            context.declarators.insert(
                (source.find('x').unwrap(), source.len() - 1),
                ScalarPolicy {
                    emitted_name: "x".into(),
                    type_hint: None,
                    zero_if_missing: false,
                    integer_wrap: None,
                },
            );
            let statement = construct_scalar_declaration(source, context, Default::default())
                .unwrap()
                .unwrap();
            let vybe_ast::StmtKind::VarDecl {
                mut declarations, ..
            } = statement.kind
            else {
                unreachable!()
            };
            let actual = declarations.pop().unwrap().init.unwrap();
            let expected = owned::scalar(rule as usize, raw);
            let expected = if source.contains("=-") {
                owned::prefix(Some(vybe_ast::UnaryOp::Neg), expected)
            } else {
                expected
            };
            assert_eq!(format!("{actual:?}"), format!("{expected:?}"), "{source}");
        }
    }

    #[test]
    fn staged_declaration_binary_initializers_preserve_precedence_and_folds() {
        let literal =
            |n| vybe_ast::Expression::new(vybe_ast::ExprKind::Lit(vybe_ast::Literal::Int(n)));
        let binary = |op, left, right| {
            vybe_ast::Expression::new(vybe_ast::ExprKind::Binary {
                op,
                left: Box::new(left),
                right: Box::new(right),
            })
        };
        for (source, expected) in [
            (
                "int x=2+3*4;",
                binary(
                    vybe_ast::BinOp::Add,
                    literal(2),
                    binary(vybe_ast::BinOp::Mul, literal(3), literal(4)),
                ),
            ),
            ("int x=1==2;", literal(0)),
            ("int x=1<<40;", literal(1_i64 << 40)),
        ] {
            let mut context = DeclarationContext::new(source);
            context.declarations.insert(
                (0, source.len()),
                DeclarationPolicy {
                    kind: vybe_ast::VarDeclKind::Var,
                    statement_span: Default::default(),
                },
            );
            context.declarators.insert(
                (4, source.len() - 1),
                ScalarPolicy {
                    emitted_name: "x".into(),
                    type_hint: None,
                    zero_if_missing: false,
                    integer_wrap: None,
                },
            );
            let statement = construct_scalar_declaration(source, context, Default::default())
                .unwrap()
                .unwrap();
            let vybe_ast::StmtKind::VarDecl {
                mut declarations, ..
            } = statement.kind
            else {
                unreachable!()
            };
            let actual = declarations.pop().unwrap().init.unwrap();
            assert_eq!(format!("{actual:?}"), format!("{expected:?}"), "{source}");
        }
    }

    #[test]
    fn owned_integer_modules_use_handoff_and_match_complete_peg_ast() {
        for expression in [
            "1+2*3-4/2%2",
            "8-3-2",
            "16/4/2",
            "1==2!=3",
            "1|2^3&4",
            "1<2<<3",
            "1<=2&&3>=2||4>5",
            "1<<31",
            "1<<32",
            "1<<40",
            "1>>63",
            "1<<64",
            "2147483648+1",
            "4294967295+1",
            "9223372036854775808+1",
            "18446744073709551615ULL+1",
            "18446744073709551616ULL+1",
            "0x80000000+1",
            "0xffffffff+1",
            "0xffffffffffffffff+1",
            "0x10000000000000000+1",
            "037777777777+1",
            "01777777777777777777777ULL+1",
            "08+1",
            "0b11111111111111111111111111111111+1",
            "0b1111111111111111111111111111111111111111111111111111111111111111+1",
            "1u+2",
            "1U+2",
            "1l+2",
            "1L+2",
            "1ll+2",
            "1LL+2",
            "1ul+2",
            "1LU+2",
            "1ull+2",
            "1LLU+2",
            "1lUl+2",
            "0xffffffffu>>1",
            "0xffffffffffffffffULL>1LL",
            "1 /* between */ + 2 * 3",
            "-2",
            "+2",
            "!2",
            "~2",
            "-~!+2",
            "1+-2",
            "-1+~2*3",
            "1+2.0",
            "0x1.8p+2+.5f",
            "'a'+'\\n'",
            "'ab'+'\\012'",
            "L'\\u00ff'+u8'\\x41'",
            "u'\\u0041'+U'\\U00000042'",
            "true==false",
        ] {
            let source = format!("#define OFFSET 2\nlong f(void) {{ return {expression}; }}");
            let reference = crate::parse_with_peg(&source).unwrap();
            OWNED_HANDOFFS.with(|count| count.set(0));
            let candidate = with_backend(Backend::Owned, || crate::walker::parse(&source)).unwrap();
            assert_eq!(
                format!("{reference:?}"),
                format!("{candidate:?}"),
                "{expression}"
            );
            assert!(
                OWNED_HANDOFFS.with(Cell::get) > 0,
                "no direct handoff: {expression}"
            );
        }
    }

    #[test]
    fn owned_recapture_keeps_successful_peg_tree_and_original_resource_errors() {
        let options = grameo::MatchOptions::default();
        let syntax = grameo::engine::parse_program(
            &native_grammar::Parser,
            Rule::binary_expression.name(),
            "+",
            options,
        )
        .unwrap_err();
        assert_eq!(syntax.kind, grameo::ParseErrorKind::Syntax);
        let tree = recapture_owned_syntax(Rule::binary_expression, "1+2", options, syntax).unwrap();
        let root = tree.into_typed_pairs::<Rule>().next().unwrap();
        assert!(
            root.into_inner().count() > 0,
            "recapture must retain contextual children"
        );

        let options = grameo::MatchOptions {
            max_steps: 1,
            ..Default::default()
        };
        let error = grameo::engine::parse_program(
            &native_grammar::Parser,
            Rule::binary_expression.name(),
            "1+2",
            options,
        )
        .unwrap_err();
        assert_ne!(error.kind, grameo::ParseErrorKind::Syntax);
        let expected = format!("{error:?}");
        let actual =
            recapture_owned_syntax(Rule::binary_expression, "1+2", options, error).unwrap_err();
        assert_eq!(expected, format!("{actual:?}"));
    }

    #[test]
    fn owned_failed_reparse_clears_cache_and_late_failure_publishes_nothing() {
        with_backend(Backend::Owned, || {
            let source = "1+2";
            let pairs =
                parse_captures(Rule::binary_expression, source, Default::default()).unwrap();
            let pair = pairs.into_iter().next().unwrap();
            assert!(take_binary(&pair).is_some());
            let result = parse_captures(
                Rule::binary_expression,
                source,
                grameo::MatchOptions {
                    max_steps: 1,
                    ..Default::default()
                },
            );
            assert!(result.is_err());
            assert!(
                take_binary(&pair).is_none(),
                "stale AST survived failed reparse"
            );

            let result = parse_captures(
                Rule::program,
                "int f(void) { return 1+2; } @",
                Default::default(),
            );
            assert!(result.is_err());
            OWNED_CACHE.with(|cache| assert!(cache.borrow().as_ref().unwrap().is_empty()));
        });
    }

    #[test]
    #[ignore = "complete saved preprocessed C module PEG/native AST parity; provide VYBE_C_FRONTEND_PREPROCESSED"]
    fn native_matching_preserves_complete_preprocessed_ast() {
        let path = std::env::var("VYBE_C_FRONTEND_PREPROCESSED")
            .expect("provide the complete saved preprocessed C source");
        let source = std::fs::read_to_string(&path).unwrap();
        let peg = crate::parse_with_peg(&source).unwrap();
        let statements = peg.body.len();
        let reference = format!("{peg:?}");
        drop(peg);
        let ((native, construction), matching) = with_native_pratt(true, ||
            profile_matching(|| profile_construction(|| crate::parse_optimized(&source))));
        let native = native.unwrap();
        assert!(construction.source_pratt_matches > 1, "{construction:?}");
        assert!(matching.rules.iter().any(|row| row.rule == Rule::frontend_unit.name()
            && row.totals.calls > 1), "{matching:?}");
        if path.to_ascii_lowercase().contains("doom") || source.contains("D_DoomMain") {
            assert!(construction.direct_body_handoffs > 0, "{construction:?}");
            assert!(construction.direct_primary_handoffs > 0, "{construction:?}");
            assert!(construction.direct_syntax_handoffs > 0, "{construction:?}");
            assert!(!matching.rules.iter().any(|row| row.rule == Rule::program.name()),
                "valid Doom input recaptured the whole program: {matching:?}");
        }
        let actual = format!("{native:?}");
        assert!(reference == actual,
            "complete PEG/native AST differs at byte {:?}; reference={} native={}",
            reference.bytes().zip(actual.bytes()).position(|(a, b)| a != b),
            reference.len(), actual.len());
        eprintln!("native_complete_ast_parity input={path} source_bytes={} statements={} ast_bytes={}",
            source.len(), statements, actual.len());
        eprintln!("native_complete_ast_construction {construction:?}");
    }

    #[test]
    fn capture_shapes_spans_and_bounded_errors_are_identical() {
        fn flatten(pairs: grameo::compat::Pairs<'_, Rule>, out: &mut Vec<(Rule, usize, usize)>) {
            for pair in pairs {
                out.push((pair.as_rule(), pair.as_span().start(), pair.as_span().end()));
                flatten(pair.into_inner(), out);
            }
        }
        for source in [
            "int main(void) { return 1 + 2 * 3; }",
            "typedef int T; int f(T x) { int T=2; return (x)+T; }",
            "struct P { int x; }; int f(struct P *p) { return p->x++; }",
            "#define N 3\nint f(void) { return N; }",
            "int f(void) { return \"a\\\\b\"[0]; }",
            "int f(void) { return ;",
            "int f(void) { return 1 + ; }",
            "int f(void) { return \"unterminated; }",
        ] {
            for max_steps in [1, 50, 1000, 10_000_000, usize::MAX] {
                let options = grameo::MatchOptions {
                    max_steps,
                    ..Default::default()
                };
                let run = |backend| {
                    with_backend(backend, || {
                        parse_captures(Rule::program, source, options)
                            .map(|pairs| {
                                let mut out = Vec::new();
                                flatten(pairs, &mut out);
                                out
                            })
                            .map_err(|error| format!("{error:?}"))
                    })
                };
                assert_eq!(
                    run(Backend::Peg),
                    run(Backend::Native),
                    "{source} budget={max_steps}"
                );
                assert_eq!(
                    run(Backend::Peg),
                    run(Backend::OwnedContextual),
                    "contextual: {source} budget={max_steps}"
                );
                assert_eq!(
                    run(Backend::Peg),
                    run(Backend::ContextualPratt),
                    "native matching plus contextual Pratt: {source} budget={max_steps}"
                );
            }
        }
    }

    #[test]
    fn contextual_pratt_native_matching_keeps_ast_and_restores_selection() {
        let source = "typedef short word; int side(int);
            int f(word *p, int i, int j) {
                word local[4] = {1,2,3,4};
                return side(p[i++]) + local[j] * sizeof(word);
            }";
        let reference = crate::parse_with_peg(source).unwrap();
        let previous = BACKEND.with(Cell::get);
        let (actual, profile) = profile_construction(|| {
            with_backend(Backend::ContextualPratt, || {
                let module = crate::walker::parse(source).unwrap();
                with_backend(Backend::Peg, || {
                    assert_eq!(BACKEND.with(Cell::get), Some(Backend::Peg));
                    assert_eq!(format!("{:?}", crate::walker::parse(source).unwrap()),
                        format!("{reference:?}"));
                });
                assert_eq!(BACKEND.with(Cell::get), Some(Backend::ContextualPratt));
                module
            })
        });
        assert_eq!(format!("{actual:?}"), format!("{reference:?}"));
        assert!(profile.contextual_pratt_handoffs > 0);
        assert_eq!(BACKEND.with(Cell::get), previous);
    }

    #[test]
    #[cfg(unix)]
    #[ignore = "same-cohort complete AST parity and paired Pratt declaration-handoff CPU"]
    fn measure_pratt_declaration_composition() {
        assert!(std::env::var_os("VYBE_C_DEBUG_TIMINGS").is_none());
        let path = std::env::var("VYBE_C_FRONTEND_PREPROCESSED").unwrap();
        let source = std::fs::read_to_string(&path).unwrap();
        let expected = format!("{:?}", with_pratt_declarations(false,
            || crate::parse_contextual_pratt(&source)).unwrap());
        let peg = format!("{:?}", crate::parse_with_peg(&source).unwrap());
        assert!(expected == peg, "baseline Pratt differs from complete PEG AST");
        drop(peg);
        let (actual, counts) = profile_construction(|| crate::parse_contextual_pratt(&source));
        let actual = actual.unwrap();
        let actual_text = format!("{actual:?}");
        assert!(expected == actual_text, "complete declaration composition AST mismatch");
        assert!(counts.declaration_handoffs > 0, "{counts:?}");
        assert_eq!(counts.mixed_regions, 0);
        eprintln!("pratt_declaration_ast_parity input={path} bytes={} statements={} {counts:?}",
            source.len(), actual.body.len());
        drop((expected, actual, actual_text));
        let mut cpus = [Vec::new(), Vec::new()];
        let mut walls = [Vec::new(), Vec::new()];
        for sample in 0..3 {
            for index in if sample % 2 == 0 { [0, 1] } else { [1, 0] } {
                let started = Instant::now();
                let cpu = cpu_ns();
                let module = with_pratt_declarations(index == 1,
                    || crate::parse_contextual_pratt(&source)).unwrap();
                drop(std::hint::black_box(module));
                let used = cpu_ns() - cpu;
                let elapsed = started.elapsed().as_nanos();
                cpus[index].push(used);
                walls[index].push(elapsed);
                eprintln!("pratt_declaration_sample enabled={} sample={sample} cpu_ns={used} wall_ns={elapsed}",
                    index == 1);
            }
        }
        for index in 0..2 {
            cpus[index].sort_unstable();
            walls[index].sort_unstable();
            eprintln!("pratt_declaration_median enabled={} cpu_ns={} wall_ns={} samples=3",
                index == 1, cpus[index][1], walls[index][1]);
        }
    }

    #[test]
    #[cfg(unix)]
    #[ignore = "complete preprocessed-Doom AST parity and three-way frontend CPU"]
    fn measure_contextual_pratt_frontend() {
        assert!(std::env::var_os("VYBE_C_DEBUG_TIMINGS").is_none(),
            "collect instrumented phases separately from CPU samples");
        let path = std::env::var("VYBE_C_FRONTEND_PREPROCESSED")
            .expect("provide the walker hook's preprocessed source dump");
        let source = std::fs::read_to_string(&path).unwrap();
        let backends = [Backend::Peg, Backend::Native, Backend::ContextualPratt];
        let controls = ["PEG", "NativeCapture", if PRATT_DECLARATIONS_DISABLED.with(Cell::get) {
            "NativeCapture+ContextualPratt(declarations=disabled)"
        } else {
            "NativeCapture+ContextualPratt(declarations=enabled)"
        }];
        eprintln!("contextual_pratt_controls input={path} controls={controls:?}; Pratt/Native includes the selected declaration-hook path, not expression construction alone");
        let reference = format!("{:?}", crate::parse_with_peg(&source).unwrap());
        for backend in &backends[1..] {
            let (actual, counts) = profile_construction(|| {
                with_backend(*backend, || crate::walker::parse(&source))
            });
            let actual = actual.unwrap();
            let direct = format!("{actual:?}");
            assert!(reference == direct,
                "{backend:?} complete AST differs at byte {:?}; reference={} direct={}",
                reference.bytes().zip(direct.bytes()).position(|(a,b)| a != b),
                reference.len(), direct.len());
            if *backend == Backend::ContextualPratt {
                assert!(counts.contextual_pratt_handoffs > 0);
            } else {
                assert_eq!(counts.contextual_pratt_handoffs, 0);
            }
            eprintln!("contextual_pratt_ast_parity {backend:?} input={path} bytes={} statements={} {counts:?}",
                source.len(), actual.body.len());
        }
        drop(reference);
        let samples = std::env::var("VYBE_C_FRONTEND_SAMPLES").ok()
            .and_then(|value| value.parse::<usize>().ok()).unwrap_or(3);
        assert!(samples >= 3);
        let mut cpus: [Vec<u128>; 3] = std::array::from_fn(|_| Vec::new());
        let mut walls: [Vec<u128>; 3] = std::array::from_fn(|_| Vec::new());
        let pairs = [(0, 1), (1, 2), (0, 2)];
        let mut cpu_ratios: [Vec<f64>; 3] = std::array::from_fn(|_| Vec::new());
        let mut wall_ratios: [Vec<f64>; 3] = std::array::from_fn(|_| Vec::new());
        // Rotate every backend through all positions; reverse the next cycle
        // to balance pair order when six or more sample rounds are requested.
        let orders = [[0, 1, 2], [1, 2, 0], [2, 0, 1],
            [2, 1, 0], [0, 2, 1], [1, 0, 2]];
        for sample in 0..samples {
            for (position, index) in orders[sample % orders.len()].into_iter().enumerate() {
                let started = std::time::Instant::now();
                let cpu = cpu_ns();
                let module = with_backend(backends[index], || crate::walker::parse(&source)).unwrap();
                drop(std::hint::black_box(module));
                let used = cpu_ns() - cpu;
                let elapsed = started.elapsed().as_nanos();
                cpus[index].push(used);
                walls[index].push(elapsed);
                eprintln!("contextual_pratt_sample {} sample={sample} position={position} cpu_ns={used} wall_ns={elapsed}", controls[index]);
            }
            for (pair, &(baseline, candidate)) in pairs.iter().enumerate() {
                assert!(cpus[baseline][sample] > 0 && walls[baseline][sample] > 0);
                let cpu = cpus[candidate][sample] as f64 / cpus[baseline][sample] as f64;
                let wall = walls[candidate][sample] as f64 / walls[baseline][sample] as f64;
                cpu_ratios[pair].push(cpu);
                wall_ratios[pair].push(wall);
                eprintln!("contextual_pratt_pair baseline={:?} candidate={:?} sample={sample} cpu_ratio={cpu:.6} wall_ratio={wall:.6}",
                    controls[baseline], controls[candidate]);
            }
        }
        for index in 0..backends.len() {
            cpus[index].sort_unstable();
            walls[index].sort_unstable();
            eprintln!("contextual_pratt_frontend {:?} input={path} cpu_median_ns={} wall_median_ns={} cpu_min_ns={} cpu_max_ns={} wall_min_ns={} wall_max_ns={} samples={samples}",
                controls[index], cpus[index][samples/2], walls[index][samples/2],
                cpus[index][0], cpus[index][samples-1], walls[index][0], walls[index][samples-1]);
        }
        for (pair, &(baseline, candidate)) in pairs.iter().enumerate() {
            cpu_ratios[pair].sort_by(f64::total_cmp);
            wall_ratios[pair].sort_by(f64::total_cmp);
            eprintln!("contextual_pratt_paired_median baseline={:?} candidate={:?} cpu_ratio={} wall_ratio={} samples={samples}",
                controls[baseline], controls[candidate],
                cpu_ratios[pair][samples / 2], wall_ratios[pair][samples / 2]);
        }
    }

    #[test]
    #[cfg(unix)]
    #[ignore = "paired ordinary-debug matching measurement"]
    fn measure_matching() {
        let root = std::path::Path::new(env!("CARGO_MANIFEST_DIR")).join("../..");
        let names = std::env::var("VYBE_C_FRONTEND_MATCH_INPUT")
            .map(|name| vec![name])
            .unwrap_or_else(|_| {
                vec![
                    "examples/c/sdlcalc/calc.c".into(),
                    "examples/c/doom_funcptr_probe.c".into(),
                ]
            });
        for name in names {
            let source = std::fs::read_to_string(root.join(&name)).unwrap();
            let options = grameo::MatchOptions::default();
            let mut times = [Vec::new(), Vec::new(), Vec::new(), Vec::new()];
            let mut cpus = [Vec::new(), Vec::new(), Vec::new(), Vec::new()];
            let backends = [
                Backend::Peg,
                Backend::Native,
                Backend::Owned,
                Backend::OwnedContextual,
            ];
            let samples = std::env::var("VYBE_C_FRONTEND_SAMPLES")
                .ok()
                .and_then(|value| value.parse::<usize>().ok())
                .unwrap_or(20);
            assert!(samples >= 3);
            for sample in 0..samples + 2 {
                for index in if sample % 2 == 0 {
                    [0, 1, 2, 3]
                } else {
                    [3, 2, 1, 0]
                } {
                    let backend = backends[index];
                    let start = std::time::Instant::now();
                    let cpu = cpu_ns();
                    let pairs =
                        with_backend(backend, || parse_captures(Rule::program, &source, options))
                            .unwrap();
                    std::hint::black_box(pairs);
                    if sample >= 2 {
                        cpus[index].push(cpu_ns() - cpu);
                        times[index].push(start.elapsed().as_nanos());
                    }
                }
            }
            for (index, values) in times.iter_mut().enumerate() {
                values.sort_unstable();
                cpus[index].sort_unstable();
                eprintln!(
                    "matching {name} {:?} wall_median_ns={} cpu_median_ns={} samples={}",
                    backends[index],
                    values[values.len() / 2],
                    cpus[index][values.len() / 2],
                    values.len()
                );
            }
            // Keep construction counts outside the uninstrumented samples.
            for backend in backends {
                let (pairs, counts) = profile_construction(|| {
                    with_backend(backend, || parse_captures(Rule::program, &source, options))
                        .unwrap()
                });
                std::hint::black_box(pairs);
                eprintln!("matching_construction {name} {backend:?} {counts:?}");
            }
        }
    }
}
