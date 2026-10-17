//! Grammar-declared source-load intent, independent of any AST or execution VM.
//! Embeddings construct requests transactionally and publish preparation work
//! only after commitment. This metadata grants no permission to execute a load.
use crate::{CompiledGrammar, Diagnostic, ExprKind, Span, grammar::RuleId};
mod collector;
pub use collector::Collector;
mod with_builder;
pub use with_builder::WithBuilder;
mod shortcuts;
pub use shortcuts::opaque_rules_preserving_requests;

/// Compiler storage/linkage policy, independent of source name-resolution scope.
/// Request space permits reuse only within the embedding's live request/VM.
/// It grants no permission to resolve or execute a load during parsing.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub enum CodeSpace { #[default] Unit, Request }
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Phase { Preprocess, Module, Execute }
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Repetition { Always, Once }
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Missing { Warning, Error, Diagnostic }
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Scope { Caller, Global, Module, Explicit }
/// Scheduling context, independent of execution scope and load policy.
/// Immediate regions may still be conditional; this never proves execution.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum PreparationContext { Immediate, Deferred }
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Policy {
    pub phase: Phase,
    pub repetition: Repetition,
    pub missing: Missing,
    pub scope: Scope,
}
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Binding<'a> {
    pub rule: RuleId,
    pub target: RuleId,
    pub selector: Option<RuleId>,
    pub loader: &'a str,
    pub policy: Policy,
}
/// Owned intent produced by a transactional frontend constructor. `T` may be
/// the embedding's common AST expression or an arena handle. Keeping the
/// unevaluated target prevents computed paths from running during parsing.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Request<T> {
    pub target: T,
    pub loader: String,
    pub policy: Policy,
    pub span: Span,
    pub preparation: PreparationContext,
}
impl Binding<'_> {
    pub fn request<T>(self, target: T, span: Span) -> Request<T> {
        Request {target, loader:self.loader.to_owned(), policy:self.policy, span, preparation:PreparationContext::Immediate}
    }
}
#[derive(Debug, Clone)]
pub(crate) struct OwnedBinding {
    pub rule: RuleId,
    pub target: RuleId,
    pub selector: Option<RuleId>,
    pub loader: String,
    pub policy: Policy,
}
impl OwnedBinding {
    pub(crate) fn borrowed(&self) -> Binding<'_> {
        Binding { rule: self.rule, target: self.target, selector: self.selector,
            loader: &self.loader, policy: self.policy }
    }
}
impl CompiledGrammar {
    pub fn source_unit_bindings(&self) -> impl Iterator<Item = Binding<'_>> {
        self.source_units.iter().map(OwnedBinding::borrowed)
    }
}

fn reachable(grammar: &CompiledGrammar, root: RuleId, wanted: RuleId) -> bool {
    let mut rules = vec![root];
    let mut seen = std::collections::HashSet::new();
    while let Some(rule) = rules.pop() {
        if !seen.insert(rule) { continue; }
        let mut pending = vec![grammar.syntax().rules[rule].expression];
        while let Some(id) = pending.pop() {
            match &grammar.syntax().expressions[id].kind {
                ExprKind::Reference(name) => if let Some(child) = grammar.rule_id(name) {
                    if child == wanted { return true; }
                    rules.push(child);
                },
                ExprKind::Sequence {left, right} | ExprKind::Choice {left, right} => {
                    pending.push(*right); pending.push(*left);
                }
                ExprKind::Repeat {expression, ..} | ExprKind::Group(expression)
                | ExprKind::Push(expression) => pending.push(*expression),
                // Lookahead does not publish semantic children.
                _ => {},
            }
        }
    }
    false
}

pub(crate) fn bind(grammar: &mut CompiledGrammar, text: &str, span: Span) -> Result<(), Vec<Diagnostic>> {
    let fail = |message: String| vec![Diagnostic::new("G016", message, span)];
    let mut words = text.split_ascii_whitespace();
    let name = words.next().ok_or_else(|| fail("expected source_unit target rule".into()))?;
    let rule = grammar.rule_id(name).ok_or_else(|| fail(format!("unknown source_unit rule '{name}'")))?;
    let mut properties = std::collections::HashMap::new();
    for word in words {
        let (key, value) = word.split_once('=').filter(|(_, v)| !v.is_empty())
            .ok_or_else(|| fail(format!("expected source_unit key=value, found '{word}'")))?;
        if !matches!(key, "target" | "selector" | "loader" | "phase" | "repeat" | "missing" | "scope") {
            return Err(fail(format!("unknown source_unit property '{key}'")));
        }
        if properties.insert(key, value).is_some() { return Err(fail(format!("duplicate source_unit property '{key}'"))); }
    }
    let get = |key| properties.get(key).copied().ok_or_else(|| fail(format!("missing source_unit property '{key}'")));
    let resolve = |name: &str| grammar.rule_id(name).ok_or_else(|| fail(format!("unknown source_unit child rule '{name}'")));
    let target = resolve(get("target")?)?;
    let selector = properties.get("selector").map(|name| resolve(name)).transpose()?;
    for child in std::iter::once(target).chain(selector) {
        if child == rule || !reachable(grammar, rule, child) {
            return Err(fail(format!("source_unit child '{}' is not a semantic descendant of '{name}'", grammar.syntax().rules[child].name)));
        }
    }
    let invalid = |key| fail(format!("invalid source_unit property '{key}'"));
    let policy = Policy {
        phase: match get("phase")? { "preprocess" => Phase::Preprocess, "module" => Phase::Module, "execute" => Phase::Execute, _ => return Err(invalid("phase")) },
        repetition: match get("repeat")? { "always" => Repetition::Always, "once" => Repetition::Once, _ => return Err(invalid("repeat")) },
        missing: match get("missing")? { "warning" => Missing::Warning, "error" => Missing::Error, "diagnostic" => Missing::Diagnostic, _ => return Err(invalid("missing")) },
        scope: match get("scope")? { "caller" => Scope::Caller, "global" => Scope::Global, "module" => Scope::Module, "explicit" => Scope::Explicit, _ => return Err(invalid("scope")) },
    };
    if grammar.source_units.iter().any(|b| b.rule == rule && b.selector == selector) {
        return Err(fail(format!("duplicate source_unit binding for '{name}' and selector")));
    }
    let loader = get("loader")?.to_owned();
    grammar.source_units.push(OwnedBinding {rule, target, selector, loader, policy});
    Ok(())
}
