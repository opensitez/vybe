//! Deterministic, build-time generation of static parser programs.
//! Generated programs require only grameo at runtime, not this crate.
use std::fmt::Write;
mod ast;
mod nullable;
mod native_tokens;
mod choice_heads;
pub use grameo::modules::{Import, Module, ModuleDiagnostic};
use grameo::{
    CompiledGrammar, Diagnostic,
    program::{Instruction, Program},
};

#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub struct GenerationOptions {
    /// Experimental failure dispatch; whole AST performance is not yet a
    /// consistent improvement, so ordinary generated parsers leave it off.
    pub literal_choice_failures: bool,
}

/// Public entry points needed by an embedding. Standalone generation retains
/// the complete API; private integrations can omit methods they never call.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct ParserApi {
    pub recognize: bool,
    pub build_ast: bool,
    pub build_captures: bool,
}
impl ParserApi {
    pub const CAPTURE_ONLY: Self = Self { recognize: false, build_ast: false, build_captures: false };
    pub const HYBRID_CAPTURES: Self = Self { build_captures: true, ..Self::CAPTURE_ONLY };
}
impl Default for ParserApi {
    fn default() -> Self { Self { recognize: true, build_ast: true, build_captures: true } }
}

pub fn generate_modules(modules: &[Module<'_>]) -> Result<String, Vec<ModuleDiagnostic>> {
    let linked = grameo::modules::compile(modules)?;
    emit(linked.grammar()).map_err(|errors| {
        errors
            .into_iter()
            .map(|error| ModuleDiagnostic {
                code: error.code,
                message: error.message,
                location: linked.location(error.span),
                related: error
                    .related
                    .into_iter()
                    .map(|(span, message)| (linked.location(span), message))
                    .collect(),
            })
            .collect()
    })
}

/// Emit a self-contained Rust module body with `Parser` and typed `Rule`.
/// Grammar validation happens before emission. The generated backend currently
/// shares the iterative matching engine; it eliminates runtime grammar loading,
/// resolution and owned grammar data, rather than specializing every matcher.
pub fn generate(source: &str) -> Result<String, Vec<Diagnostic>> {
    generate_with_options(source, GenerationOptions::default())
}

/// Explicitly enable AST annotations and emit their construction entry points.
pub fn generate_with_ast_bindings(source: &str) -> Result<String, Vec<Diagnostic>> {
    emit(&grameo::compile_with_ast_bindings(source)?)
}

/// Generate bindings only after checking the target's declared constructor types.
pub fn generate_with_ast_schema(source: &str, schema: &[grameo::ast::ConstructorSchema<'_>]) -> Result<String, Vec<Diagnostic>> {
    let grammar = grameo::compile_with_ast_bindings(source)?;
    emit_with_ast_schema(&grammar, schema, ParserApi::default())
}

pub fn generate_with_api(source: &str, api: ParserApi) -> Result<String, Vec<Diagnostic>> {
    emit_with_api(&grameo::compile(source)?, api)
}

pub fn generate_with_options(
    source: &str,
    options: GenerationOptions,
) -> Result<String, Vec<Diagnostic>> {
    let mut grammar = grameo::compile(source)?;
    grammar.set_literal_choice_failures(options.literal_choice_failures);
    emit(&grammar)
}

pub fn emit(grammar: &CompiledGrammar) -> Result<String, Vec<Diagnostic>> {
    emit_with_api(grammar, ParserApi::default())
}

pub fn emit_with_api(grammar: &CompiledGrammar, api: ParserApi) -> Result<String, Vec<Diagnostic>> {
    emit_impl(grammar, api, None, false)
}

pub fn emit_with_ast_schema(grammar: &CompiledGrammar, schema: &[grameo::ast::ConstructorSchema<'_>], api: ParserApi) -> Result<String, Vec<Diagnostic>> {
    grammar.validate_ast_schema(schema)?;
    emit_impl(grammar, api, Some(schema), false)
}

/// Explicitly specialize hybrid ordered choices and rule-call heads. Normal emission
/// and ordinary PEG capture parsing retain their existing dispatch.
pub fn emit_with_native_hybrid_choices(
    grammar: &CompiledGrammar,
    api: ParserApi,
) -> Result<String, Vec<Diagnostic>> {
    emit_impl(grammar, api, None, true)
}

fn emit_impl(grammar: &CompiledGrammar, api: ParserApi, schema: Option<&[grameo::ast::ConstructorSchema<'_>]>, native_choices: bool) -> Result<String, Vec<Diagnostic>> {
    let native_heads = native_choices || grammar.optimization_selection().native_rule_heads;
    let native_choices = native_choices || grammar.optimization_selection().native_choices;
    if !grammar.analysis().errors.is_empty() {
        return Err(grammar.analysis().errors.clone());
    }
    let errors: Vec<_> = grammar
        .syntax()
        .rules
        .iter()
        .filter(|rule| {
            matches!(
                rule.name.as_str(),
                "self" | "Self" | "super" | "crate" | "_"
            )
        })
        .map(|rule| {
            Diagnostic::new(
                "G013",
                "rule name cannot be represented as a Rust identifier",
                rule.name_span,
            )
        })
        .collect();
    if !errors.is_empty() {
        return Err(errors);
    }
    let mut out = String::from(
        "// Generated by grameo_codegen. Do not edit.\n#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]\n#[allow(non_camel_case_types)]\npub enum Rule {\n",
    );
    for rule in &grammar.syntax().rules {
        writeln!(out, "    r#{},", rule.name).unwrap();
    }
    let synthetic_eoi = grammar.rule_id("EOI").is_none();
    if synthetic_eoi {
        out.push_str("    EOI,\n");
    }
    out.push_str("}\nimpl Rule {\n    pub fn name(self) -> &'static str { match self {\n");
    for rule in &grammar.syntax().rules {
        writeln!(out, "        Self::r#{} => {:?},", rule.name, rule.name).unwrap();
    }
    if synthetic_eoi {
        out.push_str("        Self::EOI => \"EOI\",\n");
    }
    out.push_str("    }}\n    #[inline(always)]\n    pub fn from_capture(rule: ::grameo::CaptureRule) -> Option<Self> {\n        const RULES: &[Rule] = &[\n");
    for rule in &grammar.syntax().rules {
        writeln!(out, "            Rule::r#{},", rule.name).unwrap();
    }
    out.push_str("        ];\n        match rule {\n        ::grameo::CaptureRule::User(id) if id < RULES.len() => Some(RULES[id]),\n");
    out.push_str("        ::grameo::CaptureRule::EndOfInput => Some(Self::EOI),\n");
    out.push_str("        _ => None,\n    }}\n}\n#[derive(Debug, Clone, Copy, Default)]\npub struct Parser;\nimpl ::grameo::program::Program for Parser {\n    fn rule_id(&self, name: &str) -> Option<usize> { match name {\n");
    let mut entries: Vec<_> = grammar.entries().collect();
    entries.sort_unstable_by_key(|(name, _)| *name);
    for (name, id) in entries {
        writeln!(out, "        {name:?} => Some({id}),").unwrap();
    }
    out.push_str("        _ => None,\n    }}\n");
    out.push_str("    fn fast_ascii_repetitions(&self) -> bool { true }\n");
    native_tokens::emit(&mut out, grammar);
    let delimited: Vec<_> = (0..grammar.syntax().rules.len())
        .filter_map(|rule| grameo::quoted::certify_delimited(grammar, rule).map(|fragment| (rule, fragment)))
        .collect();
    if !delimited.is_empty() {
        out.push_str("    #[inline(always)]\n    fn delimited_fragment(&self, rule:usize)->Option<::grameo::quoted::DelimitedFragment<'_>>{ match rule {\n");
        for (rule, fragment) in delimited {
            writeln!(out, "        {rule} => Some(::grameo::quoted::DelimitedFragment {{ open: {:?}, close: {:?} }}),", fragment.open, fragment.close).unwrap();
        }
        out.push_str("        _ => None,\n    }}\n");
    }
    let quoted: Vec<_> = (0..grammar.syntax().rules.len())
        .filter_map(|rule| grameo::quoted::certify(grammar, rule).map(|fragment| (rule, fragment)))
        .collect();
    if !quoted.is_empty() {
        out.push_str("    #[inline(always)]\n    fn quoted_fragment(&self, rule:usize)->Option<::grameo::quoted::Fragment>{ match rule {\n");
        for (rule, fragment) in quoted {
            writeln!(out, "        {rule} => Some(::grameo::quoted::Fragment {{ delimiter: {}, escapes: {:?}, any_escape: {}, doubled_delimiter: {}, forbidden_newline: {}, newline_crlf_only: {} }}),",
                fragment.delimiter, fragment.escapes, fragment.any_escape, fragment.doubled_delimiter, fragment.forbidden_newline, fragment.newline_crlf_only).unwrap();
        }
        out.push_str("        _ => None,\n    }}\n");
    }
    out.push_str("    fn whitespace_run_class(&self, scope:usize)->Option<::grameo::lexical::ByteClass>{ match scope {\n");
    let mut classes = vec![None; grammar.syntax().expressions.len()];
    for id in 0..classes.len() {
        let mut class = grameo::lexical::ByteClass([0, 0]);
        let mut insert = |byte: u8| { class.0[usize::from(byte / 64)] |= 1u64 << (byte % 64); };
        classes[id] = match grammar.instruction(id) {
            Instruction::Literal { text, insensitive } if text.len() == 1 && text.is_ascii() => {
                let byte = text.as_bytes()[0]; insert(byte);
                if insensitive { insert(byte.to_ascii_lowercase()); insert(byte.to_ascii_uppercase()); }
                Some(class)
            },
            Instruction::Range { start, end } if start.is_ascii() && end.is_ascii() => {
                for byte in start as u8..=end as u8 { insert(byte); } Some(class)
            },
            Instruction::Group(child) => classes[child],
            Instruction::Choice { left, right } => match (classes[left], classes[right]) {
                (Some(a), Some(b)) => Some(grameo::lexical::ByteClass([a.0[0] | b.0[0], a.0[1] | b.0[1]])),
                _ => None,
            },
            _ => None,
        };
    }
    for (scope, spec) in grammar.scopes().iter().enumerate() {
        if let Some(rule) = spec.whitespace && grammar.rule(rule).mode == grameo::RuleMode::Silent {
            let (body, _) = fragment_unwrap(grammar, grammar.rule(rule).expression);
            if let Instruction::Repeat { expression, min: 1, max: None } = grammar.instruction(body)
                && let Some(class) = classes[expression] {
                writeln!(out, "        {scope}=>Some(::grameo::lexical::ByteClass({:?})),", class.0).unwrap();
            }
        }
    }
    out.push_str("        _=>None,\n    }}\n");

    let mut fragment_analysis = FragmentAnalysis::new(grammar);
    fragment_analysis.ascii_xid = grammar.optimization_selection().ascii_xid;
    let fragments = (0..grammar.syntax().rules.len())
        .map(|rule| fragment_spec(grammar, rule, &mut fragment_analysis))
        .collect::<Vec<_>>();
    let general_expressions = general_driver_expressions(grammar);
    let mut general_fragment_rules = vec![false; fragments.len()];
    for (id, &is_general) in general_expressions.iter().enumerate() {
        if is_general && let Instruction::Call(rule) = grammar.instruction(id) {
            general_fragment_rules[rule] |= fragments[rule].is_some();
        }
    }
    let has_fragments = general_fragment_rules.iter().any(|&used| used);
    if has_fragments {
        out.push_str("    fn atomic_fragments(&self)->Option<()> { Some(()) }\n");
        out.push_str("    #[inline]\n    fn atomic_fragment(&self, rule: usize)->Option<&::grameo::program::AtomicFragment<'_>> { match rule {\n");
        for (rule, &used) in general_fragment_rules.iter().enumerate() {
            if used && fragment_analysis.partial_rules.get(rule)!=Some(&true) {
                writeln!(out, "        {rule} => Some(&FRAGMENT_RULE_{rule}),").unwrap();
            }
        }
        out.push_str("        _ => None,\n    }}\n");
        if general_fragment_rules.iter().enumerate().any(|(rule,&used)| used && fragment_analysis.partial_rules.get(rule)==Some(&true)) {
            out.push_str("    #[inline]\n    fn partial_atomic_fragment(&self, rule: usize)->Option<&::grameo::program::AtomicFragment<'_>> { match rule {\n");
            for (rule,&used) in general_fragment_rules.iter().enumerate() {
                if used && fragment_analysis.partial_rules.get(rule)==Some(&true) {
                    writeln!(out, "        {rule} => Some(&FRAGMENT_RULE_{rule}),").unwrap();
                }
            }
            out.push_str("        _ => None,\n    }}\n");
        }
    }

    let mut choices = grammar.literal_choices().peekable();
    if choices.peek().is_some() {
        out.push_str("    #[inline]\n    fn literal_choice_failure(&self, expression: usize, source: &str) -> Option<usize> { match expression {\n");
        for (id, choice) in choices {
            write!(out, "        {id} => {{ if ").unwrap();
            for (index, &alternative) in choice.alternatives.iter().enumerate() {
                if index > 0 {
                    out.push_str(" || ");
                }
                let Instruction::Literal { text, insensitive } = grammar.instruction(alternative)
                else {
                    unreachable!()
                };
                write!(
                    out,
                    "::grameo::lexical::literal_matches(source, {text:?}, {insensitive})"
                )
                .unwrap();
            }
            writeln!(out, " {{ None }} else {{ Some({}) }} }},", choice.steps).unwrap();
        }
        out.push_str("        _ => None,\n    }}\n");
    }
    if native_choices || native_heads {
        out.push_str("    #[inline]\n    fn native_hybrid_choices(&self) -> bool { true }\n");
    }
    if native_choices {
        emit_native_choices(&mut out, grammar);
    }
    if native_heads {
        emit_native_rule_heads(&mut out, grammar);
    }
    out.push_str("    #[inline]\n    fn borrowed_failure_prefixes(&self) -> bool { true }\n");
    out.push_str("    fn trivia_failure_prefix_ref(&self, rule: usize) -> Option<&::grameo::lexical::FailurePrefix> { const PREFIXES: &[::grameo::lexical::FailurePrefix] = &[\n");
    let mut prefix_ids = Vec::new();
    for id in 0..grammar.syntax().rules.len() {
        if let Some(prefix) = grammar.trivia_failure_prefix(id) {
            prefix_ids.push(id);
            writeln!(out, "        ::grameo::lexical::FailurePrefix {{ bytes: {:?}, steps: {}, depth: {} }},", prefix.bytes, prefix.steps, prefix.depth).unwrap();
        }
    }
    out.push_str("    ]; match rule {\n");
    for (index, id) in prefix_ids.iter().enumerate() {
        writeln!(out, "        {id} => Some(&PREFIXES[{index}]),").unwrap();
    }
    out.push_str("        _ => None,\n    }}\n");
    if (0..grammar.syntax().rules.len()).any(|id| grammar.trivia_failure_prefix(id).is_some()) {
        out.push_str("    fn trivia_failure_prefix(&self, rule: usize) -> Option<::grameo::lexical::FailurePrefix> { match rule {\n");
        for id in 0..grammar.syntax().rules.len() {
            if let Some(prefix) = grammar.trivia_failure_prefix(id) {
                writeln!(out, "        {id} => Some(::grameo::lexical::FailurePrefix {{ bytes: {:?}, steps: {}, depth: {} }}),", prefix.bytes, prefix.steps, prefix.depth).unwrap();
            }
        }
        out.push_str("        _ => None,\n    }}\n");
    }
    out.push_str("    #[inline]\n    fn trivia_failure_cost(&self, scope: usize, whitespace_done: bool, byte: Option<u8>) -> ::grameo::program::TriviaFailure { match (scope, whitespace_done) {\n");
    for (id, scope) in grammar.scopes().iter().enumerate() {
        for whitespace_done in [false, true] {
            let mut bytes = [0u64; 4];
            let mut cost = Some(if whitespace_done { 1usize } else { 2 });
            let mut depth = 0usize;
            for rule in scope
                .whitespace
                .filter(|_| !whitespace_done)
                .into_iter()
                .chain(scope.comment)
            {
                let Some(prefix) = grammar.trivia_failure_prefix(rule) else {
                    cost = None;
                    break;
                };
                cost = cost.and_then(|cost| cost.checked_add(prefix.steps)?.checked_add(2));
                let Some(next_depth) = prefix.depth.checked_add(1) else {
                    cost = None;
                    break;
                };
                depth = depth.max(next_depth);
                for (slot, bits) in bytes.iter_mut().zip(prefix.bytes) {
                    *slot |= bits;
                }
            }
            if let Some(cost) = cost {
                writeln!(out, "        ({id}, {whitespace_done}) => if byte.is_none_or(|byte| {:?}[usize::from(byte / 64)] & (1u64 << (byte % 64)) == 0) {{ ::grameo::program::TriviaFailure::Miss {{ steps: {cost}, depth: {depth} }} }} else {{ ::grameo::program::TriviaFailure::Possible }},", bytes).unwrap();
            }
        }
    }
    out.push_str("        _ => ::grameo::program::TriviaFailure::Unknown,\n    }}\n");
    out.push_str("    #[inline]\n    fn leading_bytes(&self, expression: usize) -> Option<::grameo::lexical::LeadingBytes> { LEADING_BYTES[expression] }\n");
    out.push_str("    #[inline]\n    fn atomic_expression(&self, expression: usize) -> bool { ATOMIC_EXPRESSIONS[expression] }\n");
    let mut terminal_bitmap = vec![0u64; grammar.syntax().expressions.len().div_ceil(64)];
    for id in 0..grammar.syntax().expressions.len() {
        if grammar.terminal_choice(id).is_some() {
            terminal_bitmap[id / 64] |= 1u64 << (id % 64);
        }
    }
    writeln!(out, "    #[inline]\n    fn has_terminal_choice(&self, expression: usize) -> bool {{ const BITS: &[u64] = &{:?}; BITS.get(expression / 64).is_some_and(|word| word & (1u64 << (expression % 64)) != 0) }}", terminal_bitmap).unwrap();
    out.push_str("    fn terminal_choice(&self, expression: usize) -> Option<::grameo::lexical::TerminalChoice<'_>> { match expression {\n");
    for id in 0..grammar.syntax().expressions.len() {
        if let Some(choice) = grammar.terminal_choice(id) {
            write!(out, "        {id} => Some(::grameo::lexical::TerminalChoice {{ failure_steps: {}, alternatives: &[", choice.failure_steps).unwrap();
            for alternative in choice.alternatives {
                write!(out, "::grameo::lexical::TerminalAlternative {{ expression: {}, success_steps: {} }},", alternative.expression, alternative.success_steps).unwrap();
            }
            out.push_str("] }),\n");
        }
    }
    out.push_str("        _ => None,\n    }}\n");
    out.push_str("    #[inline]\n    fn terminal_candidates(&self, expression: usize, _byte: Option<u8>) -> Option<&[usize]> { match expression {\n");
    for id in 0..grammar.syntax().expressions.len() {
        let Some(choice) = grammar.terminal_choice(id) else {
            continue;
        };
        if !choice.alternatives.iter().all(|a| matches!(grammar.instruction(a.expression), Instruction::Literal { text, .. } if !text.is_empty())) { continue; }
        let mut groups = std::collections::BTreeMap::<Vec<usize>, Vec<u8>>::new();
        // Derive first-byte buckets once per alternative, rather than rescanning
        // every alternative for each of the 256 input bytes. Iteration order
        // preserves the exact original candidate priority and emitted grouping.
        let mut buckets: [Vec<usize>; 256] = std::array::from_fn(|_| Vec::new());
        for (index, alternative) in choice.alternatives.iter().enumerate() {
            let Instruction::Literal { text, insensitive } =
                grammar.instruction(alternative.expression)
            else {
                unreachable!()
            };
            let first = text.as_bytes()[0];
            buckets[usize::from(first)].push(index);
            if insensitive && first.is_ascii_alphabetic() {
                let other = if first.is_ascii_lowercase() {
                    first.to_ascii_uppercase()
                } else {
                    first.to_ascii_lowercase()
                };
                buckets[usize::from(other)].push(index);
            }
        }
        for (byte, indices) in (0u8..=255).zip(buckets) {
            if !indices.is_empty() {
                groups.entry(indices).or_default().push(byte);
            }
        }
        writeln!(out, "        {id} => Some(match _byte {{").unwrap();
        for (indices, bytes) in groups {
            let patterns = bytes
                .iter()
                .map(|b| format!("Some({b})"))
                .collect::<Vec<_>>()
                .join(" | ");
            writeln!(out, "            {patterns} => &{indices:?},").unwrap();
        }
        out.push_str("            _ => &[],\n        }),\n");
    }
    out.push_str("        _ => None,\n    }}\n");
    nullable::emit(&mut out, grammar);
    if grammar.native_token_compilation() && !out.contains("fn match_observers_are_noop") {out.push_str("    fn match_observers_are_noop(&self)->bool{true}\n");}
    emit_source_units(&mut out, grammar);
    let empty_heads:Vec<_>=(0..grammar.syntax().rules.len()).filter_map(|r|grammar.empty_success_head(r).map(|h|(r,*h))).collect();
    if !empty_heads.is_empty(){
        if !out.contains("fn match_observers_are_noop"){out.push_str("    fn match_observers_are_noop(&self)->bool{true}\n");}
        out.push_str("    #[inline(always)]\n    fn empty_success_candidate(&self,rule:usize)->bool{matches!(rule,");
        out.push_str(&empty_heads.iter().map(|(r,_)|r.to_string()).collect::<Vec<_>>().join("|"));out.push_str(")}\n");
        out.push_str("    fn empty_success_metadata_is_current(&self)->bool{true}\n");
        out.push_str("    fn empty_success_head(&self,rule:usize)->Option<&::grameo::program::EmptySuccessHead>{match rule{\n");
        for (r,h) in empty_heads{writeln!(out,"        {r}=>Some(&::grameo::program::EmptySuccessHead{{expression:{},scope:{},first:{:?}}}),",h.expression,h.scope,h.first).unwrap();}
        out.push_str("        _=>None,\n    }}\n");
    }

    let empty_suffixes:Vec<_>=(0..grammar.syntax().expressions.len()).filter_map(|id|grammar.empty_suffix_head(id).map(|h|(id,*h))).collect();
    if !empty_suffixes.is_empty(){
        if !out.contains("fn match_observers_are_noop"){out.push_str("    fn match_observers_are_noop(&self)->bool{true}\n");}
        out.push_str("    #[inline(always)]\n    fn empty_suffix_candidate(&self,id:usize)->bool{matches!(id,");
        out.push_str(&empty_suffixes.iter().map(|(id,_)|id.to_string()).collect::<Vec<_>>().join("|"));out.push_str(")}\n");
        out.push_str("    fn empty_suffix_metadata_is_current(&self)->bool{true}\n");
        out.push_str("    fn empty_suffix_head(&self,id:usize)->Option<&::grameo::program::EmptySuffixHead>{match id{\n");
        for (id,h)in empty_suffixes{writeln!(out,"        {id}=>Some(&::grameo::program::EmptySuffixHead{{owner:{},owner_body:{},repetition:{},child:{},scope:{},first:{:?},calls:{:?},calls_len:{},depth:{}}}),",h.owner,h.owner_body,h.repetition,h.child,h.scope,h.first,h.calls,h.calls_len,h.depth).unwrap();}
        out.push_str("        _=>None,\n    }}\n");
    }
    let owned:Vec<_>=(0..grammar.syntax().rules.len()).filter_map(|r|grammar.owned_capture_boundary(r).map(|c|(r,c))).collect();
    if !owned.is_empty(){
        if !out.contains("fn match_observers_are_noop"){out.push_str("    fn match_observers_are_noop(&self)->bool{true}\n");}
        out.push_str("    fn owned_capture_boundary_metadata_is_current(&self)->bool{true}\n");
        out.push_str("    fn owned_capture_boundary(&self,rule:usize)->Option<::grameo::program::OwnedCaptureBoundary<'_>>{match rule {\n");
        for (rule,c) in owned {writeln!(out,"        {rule}=>Some(::grameo::program::OwnedCaptureBoundary{{expression:{},scope:{},capture_rules:&{:?}}}),",c.expression,c.scope,c.capture_rules).unwrap();}
        out.push_str("        _=>None,\n    }}\n");
    }

    if !grammar.choice_head_rules().is_empty(){choice_heads::emit(&mut out,grammar);}
    out.push_str("    fn failure_terminals(&self, rule: usize) -> &[usize] { match rule {\n");
    for id in 0..grammar.syntax().rules.len() {
        let terminals = grammar.failure_terminals(id);
        if !terminals.is_empty() {
            writeln!(out, "        {id} => &{terminals:?},").unwrap();
        }
    }
    out.push_str("        _ => &[],\n    }}\n");
    if (0..grammar.syntax().expressions.len()).any(|id| grammar.shared_prefix_choice(id).is_some()) {
        if !out.contains("    fn match_observers_are_noop(") { out.push_str("    fn match_observers_are_noop(&self) -> bool { true }\n"); }
        out.push_str("    fn shared_prefix_choice(&self, id: usize) -> Option<::grameo::program::SharedPrefixChoice> { match id {\n");
        for id in 0..grammar.syntax().expressions.len() {
            if let Some(s) = grammar.shared_prefix_choice(id) {
                writeln!(out,"        {id} => Some(::grameo::program::SharedPrefixChoice {{ target:{}, target_expression:{}, branch_expression:{}, branch_sequence:{}, left_choice:{}, wrapper_call:{}, right:{}, scope:{}, prefix_rule:{}, suffix_rule:{:?}, prior:{}, wrapper:{}, prefix:{}, suffix:{} }}),",s.target,s.target_expression,s.branch_expression,s.branch_sequence,s.left_choice,s.wrapper_call,s.right,s.scope,s.prefix_rule,s.suffix_rule,s.prior,s.wrapper,s.prefix,s.suffix).unwrap();
            }
        }
        out.push_str("        _ => None,\n    }}\n");
    }
    if (0..grammar.syntax().expressions.len()).any(|id| grammar.shared_wrapper_prefix(id).is_some()) {
        if !out.contains("    fn match_observers_are_noop(") { out.push_str("    fn match_observers_are_noop(&self) -> bool { true }\n"); }
        out.push_str("    fn shared_wrapper_prefix_metadata_is_current(&self)->bool{true}\n");
        out.push_str("    fn shared_wrapper_prefix(&self,id:usize)->Option<::grameo::program::SharedWrapperPrefix>{match id {\n");
        for id in 0..grammar.syntax().expressions.len() {
            if let Some(s)=grammar.shared_wrapper_prefix(id) {
                writeln!(out,"        {id} => Some(::grameo::program::SharedWrapperPrefix {{choice:{},branch:{},fallback_rule:{},prefix:{},suffix:{},fallback:{}}}),",s.choice,s.branch,s.fallback_rule,s.prefix,s.suffix,s.fallback).unwrap();
            }
        }
        out.push_str("        _=>None,\n    }}\n");
    }
    if (0..grammar.syntax().expressions.len()).any(|id|grammar.shared_group_prefix(id).is_some()) {
        if !out.contains("fn match_observers_are_noop") { out.push_str("    fn match_observers_are_noop(&self)->bool{true}\n"); }
        out.push_str("    fn shared_group_prefix(&self,id:usize)->Option<::grameo::program::SharedGroupPrefix>{match id {\n");
        for id in 0..grammar.syntax().expressions.len(){if let Some(s)=grammar.shared_group_prefix(id){
            writeln!(out,"        {id} => Some(::grameo::program::SharedGroupPrefix {{ target:{}, target_expression:{}, choice:{}, left:{}, right:{}, left_inner:{}, open:{}, branch_call:{}, close:{}, right_inner:{}, fallback_open:{}, optional:{}, fallback_close:{}, branch:{}, branch_expression:{}, fallback_rule:{}, fallback_expression:{}, prefix_rule:{}, prefix:{}, branch_tail:{}, fallback_inner:{}, fallback_prefix:{}, commas:{}, trailing:{} }}),",s.target,s.target_expression,s.choice,s.left,s.right,s.left_inner,s.open,s.branch_call,s.close,s.right_inner,s.fallback_open,s.optional,s.fallback_close,s.branch,s.branch_expression,s.fallback_rule,s.fallback_expression,s.prefix_rule,s.prefix,s.branch_tail,s.fallback_inner,s.fallback_prefix,s.commas,s.trailing).unwrap();
        }}
        out.push_str("        _=>None,\n    }}\n");
    }
    if (0..grammar.syntax().expressions.len()).any(|id|grammar.shared_inner_prefix(id).is_some()) {
        if !out.contains("fn match_observers_are_noop") { out.push_str("    fn match_observers_are_noop(&self)->bool{true}\n"); }
        out.push_str("    fn shared_inner_prefix(&self,id:usize)->Option<::grameo::program::SharedInnerPrefix>{match id {\n");
        for id in 0..grammar.syntax().expressions.len(){if let Some(s)=grammar.shared_inner_prefix(id){
            writeln!(out,"        {id} => Some(::grameo::program::SharedInnerPrefix {{ target:{}, target_expression:{}, choice:{}, left:{}, right:{}, prefix_rule:{}, prefix:{}, first:{}, second:{}, third:{}, kind:{} }}),",s.target,s.target_expression,s.choice,s.left,s.right,s.prefix_rule,s.prefix,s.first,s.second,s.third,s.kind).unwrap();
        }}
        out.push_str("        _=>None,\n    }}\n");
    }
    let operator_rhs:Vec<_>=(0..grammar.syntax().rules.len()).flat_map(|target|grammar.source_pratt(target).into_iter().flat_map(move |spec|spec.operators.iter().filter_map(move |op|grammar.source_pratt_operator_rhs_atom(target,op.rule).map(|atom|(target,op.rule,atom))))).collect();
    if !operator_rhs.is_empty() {
        out.push_str("    fn source_pratt_operator_rhs_atom(&self, target: usize, operator: usize) -> Option<usize> { match (target, operator) {\n");
        for (target,operator,atom) in operator_rhs {writeln!(out,"        ({target}, {operator}) => Some({atom}),").unwrap();}
        out.push_str("        _ => None,\n    }}\n");
    }
    if (0..grammar.syntax().rules.len()).any(|id| grammar.source_pratt_rhs_operator(id).is_some()) {
        out.push_str("    fn source_pratt_rhs_operator(&self, rule: usize) -> Option<usize> { match rule {\n");
        for id in 0..grammar.syntax().rules.len() { if let Some(operator)=grammar.source_pratt_rhs_operator(id) {writeln!(out,"        {id} => Some({operator}),").unwrap();} }
        out.push_str("        _ => None,\n    }}\n");
    }
    if (0..grammar.syntax().rules.len()).any(|id| grammar.source_pratt_rhs_atom(id).is_some()) {
        out.push_str("    fn source_pratt_rhs_atom(&self, rule: usize) -> Option<usize> { match rule {\n");
        for id in 0..grammar.syntax().rules.len() {
            if let Some(rhs) = grammar.source_pratt_rhs_atom(id) {
                writeln!(out, "        {id} => Some({rhs}),").unwrap();
            }
        }
        out.push_str("        _ => None,\n    }}\n");
    }
    if (0..grammar.syntax().rules.len()).any(|id|grammar.source_pratt_operator_trivia(id).is_some()) {
        out.push_str("    fn source_pratt_operator_trivia(&self, rule: usize) -> Option<usize> { match rule {\n");
        for id in 0..grammar.syntax().rules.len() {
            if let Some(trivia)=grammar.source_pratt_operator_trivia(id) {writeln!(out,"        {id} => Some({trivia}),").unwrap();}
        }
        out.push_str("        _ => None,\n    }}\n");
    }
    if (0..grammar.syntax().rules.len()).any(|id| grammar.source_pratt(id).is_some()) {
        out.push_str("    fn source_pratt(&self, rule: usize) -> Option<::grameo::program::SourcePratt<'_>> { match rule {\n");
        for id in 0..grammar.syntax().rules.len() {
            if let Some(spec) = grammar.source_pratt(id) {
                writeln!(out, "        {id} => Some(::grameo::program::SourcePratt {{ atom: {}, trailing_trivia: ::grameo::islands::TrailingTrivia::{:?}, operators: &[", spec.atom, spec.trailing_trivia).unwrap();
                for operator in spec.operators {
                    writeln!(out, "            ::grameo::program::BoundOperator {{ rule: {}, precedence: {}, fixity: ::grameo::pratt::Fixity::{:?} }},", operator.rule, operator.precedence, operator.fixity).unwrap();
                }
                out.push_str("        ] }),\n");
            }
        }
        out.push_str("        _ => None,\n    }}\n");
    }
    if let Some(class) = grammar.whitespace_class() {
        writeln!(out, "    fn whitespace_class(&self) -> Option<::grameo::lexical::ByteClass> {{ Some(::grameo::lexical::ByteClass({:?})) }}", class.0).unwrap();
    }
    if let Some(class) = grammar.whitespace_prefix_class() {
        writeln!(out, "    fn whitespace_prefix_class(&self) -> Option<::grameo::lexical::ByteClass> {{ Some(::grameo::lexical::ByteClass({:?})) }}", class.0).unwrap();
    }
    let stack_free = (0..grammar.syntax().expressions.len()).all(|id| {
        !matches!(
            grammar.instruction(id),
            Instruction::Push(_)
                | Instruction::PushLiteral(_)
                | Instruction::PeekSlice { .. }
                | Instruction::Builtin(
                    grameo::Builtin::Peek
                        | grameo::Builtin::PeekAll
                        | grameo::Builtin::Pop
                        | grameo::Builtin::PopAll
                        | grameo::Builtin::Drop
                )
        )
    });
    let capture_program = if stack_free {
        "parse_stackfree_compact_program"
    } else {
        "parse_compact_program"
    };
    writeln!(out, "    #[inline]\n    fn capture_parse<'g, 'i>(&'g self, entry: &str, source: &'i str, options: ::grameo::MatchOptions) -> Result<::grameo::ParseTree<'g, 'i>, ::grameo::ParseError> {{ ::grameo::engine::{capture_program}(self, entry, source, options) }}").unwrap();
    out.push_str("    #[inline]\n    fn compact_capture_scope(&self, id: usize) -> ::grameo::program::CompactScope { COMPACT_SCOPES[id] }\n    #[inline]\n    fn scope(&self, id: usize) -> ::grameo::program::Scope { SCOPES[id] }\n    #[inline]\n    fn rule(&self, id: usize) -> ::grameo::program::RuleSpec<'_> { RULES[id] }\n    #[inline]\n    fn instruction(&self, id: usize) -> ::grameo::program::Instruction<'_> { INSTRUCTIONS[id] }\n}\n");

    write!(
        out,
        "static LEADING_BYTES: [Option<::grameo::lexical::LeadingBytes>; {}] = [",
        grammar.syntax().expressions.len()
    )
    .unwrap();
    for id in 0..grammar.syntax().expressions.len() {
        match grammar.leading_bytes(id) {
            Some(p) => write!(out, "Some(::grameo::lexical::LeadingBytes {{ class: ::grameo::lexical::ByteClass({:?}), steps: {} }}),", p.class.0, p.steps).unwrap(),
            None => out.push_str("None,"),
        }
    }
    out.push_str("];\n");
    write!(
        out,
        "static ATOMIC_EXPRESSIONS: [bool; {}] = [",
        grammar.syntax().expressions.len()
    )
    .unwrap();
    for id in 0..grammar.syntax().expressions.len() {
        write!(out, "{},", grammar.atomic_expression(id)).unwrap();
    }
    out.push_str("];\n");
    writeln!(
        out,
        "static RULES: [::grameo::program::RuleSpec<'static>; {}] = [",
        grammar.syntax().rules.len()
    )
    .unwrap();
    for (id, rule) in grammar.syntax().rules.iter().enumerate() {
        writeln!(out, "    ::grameo::program::RuleSpec {{ name: {:?}, mode: ::grameo::RuleMode::{:?}, expression: {}, scope: {} }},", rule.name, rule.mode, rule.expression, grammar.rule(id).scope).unwrap();
    }
    out.push_str("];\n");
    writeln!(
        out,
        "static COMPACT_SCOPES: [::grameo::program::CompactScope; {}] = [",
        grammar.scopes().len()
    )
    .unwrap();
    for scope in 0..grammar.scopes().len() {
        writeln!(out, "::grameo::program::CompactScope::new({scope}),").unwrap();
    }
    out.push_str("];\n");
    writeln!(
        out,
        "static SCOPES: [::grameo::program::Scope; {}] = [",
        grammar.scopes().len()
    )
    .unwrap();
    for scope in grammar.scopes() {
        writeln!(out, "    ::grameo::program::Scope {{ whitespace: {:?}, comment: {:?}, whitespace_class: {}, whitespace_prefix_class: {} }},", scope.whitespace, scope.comment, byte_class(scope.whitespace_class), byte_class(scope.whitespace_prefix_class)).unwrap();
    }
    out.push_str("];\n");
    writeln!(
        out,
        "static INSTRUCTIONS: [::grameo::program::Instruction<'static>; {}] = {{",
        grammar.syntax().expressions.len()
    )
    .unwrap();
    out.push_str("    use ::grameo::program::Instruction::*;\n    [\n");
    for id in 0..grammar.syntax().expressions.len() {
        out.push_str("        ");
        write_instruction(&mut out, grammar.instruction(id)).unwrap();
        out.push_str(",\n");
    }
    out.push_str("    ]\n};\n");
    if has_fragments {
        for (rule, spec) in fragments.iter().enumerate() {
            if let Some(spec) = spec && general_fragment_rules[rule] {
                writeln!(out, "static FRAGMENT_RULE_{rule}: ::grameo::program::AtomicFragment<'static> = {spec};").unwrap();
            }
        }
    }
    out.push_str("impl Parser {\n    pub fn parse<'g, 'i>(&'g self, rule: Rule, input: &'i str) -> Result<::grameo::ParseTree<'g, 'i>, ::grameo::ParseError> {\n        ::grameo::engine::parse_program(self, rule.name(), input, ::grameo::MatchOptions::default())\n    }\n}\n");
    if api.recognize {
        out.push_str("impl Parser {\n    pub fn recognize(&self, rule: Rule, input: &str) -> Result<::grameo::Recognition, ::grameo::ParseError> {\n        ::grameo::engine::recognize_program(self, rule.name(), input, ::grameo::MatchOptions::default())\n    }\n}\n");
    }
    if api.build_ast {
        out.push_str("impl Parser {\n    pub fn build_ast<B: ::grameo::builder::Builder>(&self, rule: Rule, input: &str, builder: &mut B) -> Result<usize, ::grameo::ParseError> {\n        ::grameo::engine::build_ast_program(self, rule.name(), input, ::grameo::MatchOptions::default(), builder)\n    }\n}\n");
    }
    if api.build_ast || api.build_captures {
        ast::emit(grammar, &mut out, api.build_ast, schema)?;
    }
    let semantic_capture_program = if stack_free { "build_stackfree_compact_capture_program" } else { "build_compact_capture_program" };
    if api.build_captures {
        writeln!(out, "impl Parser {{\n    pub fn build_captures<'g, 'i, B: ::grameo::builder::Builder>(&'g self, rule: Rule, input: &'i str, builder: &mut B) -> Result<::grameo::ParseTree<'g, 'i>, ::grameo::ParseError> {{\n        ::grameo::engine::{semantic_capture_program}(self, rule.name(), input, ::grameo::MatchOptions::default(), builder)\n    }}\n}}\n").unwrap();
        let semantic_capture_program_at = if stack_free { "build_stackfree_compact_capture_program_at" } else { "build_compact_capture_program_at" };
        writeln!(out, "impl Parser {{\n    pub fn build_captures_at<'g, 'i, B: ::grameo::builder::Builder>(&'g self, rule: Rule, input: &'i str, start_offset: usize, builder: &mut B) -> Result<::grameo::ParseTree<'g, 'i>, ::grameo::ParseError> {{\n        ::grameo::engine::{semantic_capture_program_at}(self, rule.name(), input, start_offset, ::grameo::MatchOptions::default(), builder)\n    }}\n}}\n").unwrap();
    }
    out.push_str("impl ::grameo::compat::RuleIdentity for Rule {\n    fn from_capture(capture: ::grameo::CaptureRule) -> Option<Self> { Self::from_capture(capture) }\n}\nimpl Parser {\n    pub fn parse_pairs<'i>(&self, rule: Rule, input: &'i str) -> Result<::grameo::compat::Pairs<'i, Rule>, ::grameo::ParseError> {\n        self.parse(rule, input).map(|tree| tree.into_typed_pairs())\n    }\n}\n");
    Ok(out)
}

/// The compact atomic driver executes certified declared Atomic bodies itself.
/// Their intraprocedural call sites cannot reach the general fragment entry path.
/// Union ownership preserves shared expressions used by any general body, and
/// unknown/other modes stay eligible conservatively. Calls do not traverse into
/// callees: each rule's own mode selects its body driver independently.
fn general_driver_expressions(grammar: &CompiledGrammar) -> Vec<bool> {
    let mut general = vec![false; grammar.syntax().expressions.len()];
    for rule in 0..grammar.syntax().rules.len() {
        let definition = grammar.rule(rule);
        if definition.mode == grameo::grammar::RuleMode::Atomic
            && grammar.atomic_expression(definition.expression)
        {
            continue;
        }
        let mut pending = vec![definition.expression];
        while let Some(id) = pending.pop() {
            if std::mem::replace(&mut general[id], true) {
                continue;
            }
            match grammar.instruction(id) {
                Instruction::Sequence { left, right } | Instruction::Choice { left, right } => {
                    pending.push(left);
                    pending.push(right);
                }
                Instruction::Group(child)
                | Instruction::Push(child)
                | Instruction::Repeat {
                    expression: child, ..
                }
                | Instruction::Predicate {
                    expression: child, ..
                } => pending.push(child),
                _ => {}
            }
        }
    }
    general
}

fn fragment_unwrap(g: &CompiledGrammar, mut id: usize) -> (usize, usize) {
    let mut groups = 0;
    while let Instruction::Group(child) = g.instruction(id) {
        id = child;
        groups += 1;
    }
    (id, groups)
}
// Byte-table probes consume only match/cost. The EOF probe also builds the
// diagnostic list, preserving its original alternative order.
fn fragment_class(
    g: &CompiledGrammar,
    id: usize,
    byte: Option<u8>,
) -> Option<(bool, usize, Vec<String>)> {
    if g.terminal_choice(id).is_some() {
        return None;
    }
    let (matched, cost, expected) = match g.instruction(id) {
        Instruction::Group(child) => {
            let (m, c, e) = fragment_class(g, child, byte)?;
            (m, c + 1, e)
        }
        Instruction::Choice { left, right } => {
            let (m, c, mut e) = fragment_class(g, left, byte)?;
            if m {
                (true, c + 2, e)
            } else {
                let (m2, c2, e2) = fragment_class(g, right, byte)?;
                e.extend(e2);
                (m2, c + c2 + 2, e)
            }
        }
        Instruction::Builtin(b) => {
            use grameo::Builtin;
            let accepts: fn(u8) -> bool = match b {
                Builtin::Ascii => |b| b < 128,
                Builtin::AsciiDigit => |b| b.is_ascii_digit(),
                Builtin::AsciiAlpha => |b| b.is_ascii_alphabetic(),
                Builtin::AsciiAlphaLower => |b| b.is_ascii_lowercase(),
                Builtin::AsciiAlphaUpper => |b| b.is_ascii_uppercase(),
                Builtin::AsciiAlphanumeric => |b| b.is_ascii_alphanumeric(),
                Builtin::AsciiHexDigit => |b| b.is_ascii_hexdigit(),
                Builtin::AsciiNonzeroDigit => |b| (b'1'..=b'9').contains(&b),
                Builtin::AsciiBinDigit => |b| matches!(b, b'0' | b'1'),
                Builtin::AsciiOctDigit => |b| (b'0'..=b'7').contains(&b),
                Builtin::XidStart if g.optimization_selection().ascii_xid => |b| b.is_ascii_alphabetic(),
                Builtin::XidContinue if g.optimization_selection().ascii_xid => |b| b.is_ascii_alphanumeric() || b == b'_',
                _ => return None,
            };
            (
                byte.is_some_and(accepts),
                1,
                if byte.is_none() {
                    vec![format!(
                        "::grameo::program::FragmentExpected::Builtin(::grameo::Builtin::{b:?})"
                    )]
                } else {
                    Vec::new()
                },
            )
        }
        Instruction::Literal {
            text,
            insensitive: false,
        } if text.len() == 1 && text.is_ascii() => (
            byte == Some(text.as_bytes()[0]),
            1,
            if byte.is_none() {
                vec![format!(
                    "::grameo::program::FragmentExpected::Literal({text:?})"
                )]
            } else {
                Vec::new()
            },
        ),
        Instruction::Range { start, end } if start.is_ascii() && end.is_ascii() => (
            byte.is_some_and(|b| b >= start as u8 && b <= end as u8),
            1,
            if byte.is_none() {
                vec![format!(
                    "::grameo::program::FragmentExpected::Range({start:?},{end:?})"
                )]
            } else {
                Vec::new()
            },
        ),
        _ => return None,
    };
    Some((matched, cost, expected))
}

// Build byte lanes together, retaining unsupported lanes independently. EOF
// diagnostics continue through the original ordered recursive probe below.
struct FragmentByteCertificate {
    valid: [bool; 256],
    matched: [bool; 256],
    cost: [usize; 256],
}
impl FragmentByteCertificate {
    fn probe(&self, byte: u8) -> Option<(bool, usize, Vec<String>)> {
        let lane = usize::from(byte);
        self.valid[lane].then(|| (self.matched[lane], self.cost[lane], Vec::new()))
    }
}
impl Default for FragmentByteCertificate {
    fn default() -> Self {
        Self {
            valid: [false; 256],
            matched: [false; 256],
            cost: [0; 256],
        }
    }
}
enum FragmentMemo {
    Unknown,
    Unsupported,
    Supported(std::rc::Rc<FragmentByteCertificate>),
}
struct FragmentAnalysis {
    bytes: Vec<FragmentMemo>,
    ascii_xid: bool,
    partial_rules: Vec<bool>,
}
impl FragmentAnalysis {
    fn new(_g: &CompiledGrammar) -> Self {
        Self { bytes: Vec::new(), ascii_xid: false, partial_rules: Vec::new() }
    }
    fn eof(&self, g: &CompiledGrammar, id: usize) -> Option<(bool, usize, Vec<String>)> {
        fragment_class(g, id, None)
    }
    #[cfg(test)]
    fn probe(
        &mut self,
        g: &CompiledGrammar,
        id: usize,
        byte: u8,
    ) -> Option<(bool, usize, Vec<String>)> {
        self.certificate(g, id)?.probe(byte)
    }
    fn certificate(
        &mut self,
        g: &CompiledGrammar,
        id: usize,
    ) -> Option<std::rc::Rc<FragmentByteCertificate>> {
        if self.bytes.is_empty() {
            self.bytes
                .resize_with(g.syntax().expressions.len(), || FragmentMemo::Unknown);
        }
        match &self.bytes[id] {
            FragmentMemo::Supported(certificate) => return Some(std::rc::Rc::clone(certificate)),
            FragmentMemo::Unsupported => return None,
            FragmentMemo::Unknown => {}
        }
        if g.terminal_choice(id).is_some() {
            self.bytes[id] = FragmentMemo::Unsupported;
            return None;
        }
        let instruction = g.instruction(id);
        let supported = match instruction {
            Instruction::Group(_) | Instruction::Choice { .. } => true,
            Instruction::Builtin(b) => matches!(
                b,
                grameo::Builtin::Ascii
                    | grameo::Builtin::AsciiDigit
                    | grameo::Builtin::AsciiAlpha
                    | grameo::Builtin::AsciiAlphaLower
                    | grameo::Builtin::AsciiAlphaUpper
                    | grameo::Builtin::AsciiAlphanumeric
                    | grameo::Builtin::AsciiHexDigit
                    | grameo::Builtin::AsciiNonzeroDigit
                    | grameo::Builtin::AsciiBinDigit
                    | grameo::Builtin::AsciiOctDigit
            ) || self.ascii_xid && matches!(b, grameo::Builtin::XidStart | grameo::Builtin::XidContinue),
            Instruction::Literal {
                text,
                insensitive: false,
            } => text.len() == 1 && text.is_ascii(),
            Instruction::Range { start, end } => start.is_ascii() && end.is_ascii(),
            _ => false,
        };
        if !supported {
            self.bytes[id] = FragmentMemo::Unsupported;
            return None;
        }
        let mut certificate = FragmentByteCertificate::default();
        {
            match instruction {
                Instruction::Group(child) => {
                    let Some(child) = self.certificate(g, child) else {
                        self.bytes[id] = FragmentMemo::Unsupported;
                        return None;
                    };
                    for lane in 0..256 {
                        if child.valid[lane] {
                            certificate.valid[lane] = true;
                            certificate.matched[lane] = child.matched[lane];
                            certificate.cost[lane] = child.cost[lane] + 1;
                        }
                    }
                }
                Instruction::Choice { left, right } => {
                    let Some(left) = self.certificate(g, left) else {
                        self.bytes[id] = FragmentMemo::Unsupported;
                        return None;
                    };
                    // A right branch is required only for supported left failures.
                    let right = left
                        .valid
                        .iter()
                        .zip(&left.matched)
                        .any(|(&valid, &matched)| valid && !matched)
                        .then(|| self.certificate(g, right))
                        .flatten();
                    for lane in 0..256 {
                        if !left.valid[lane] {
                            continue;
                        }
                        if left.matched[lane] {
                            certificate.valid[lane] = true;
                            certificate.matched[lane] = true;
                            certificate.cost[lane] = left.cost[lane] + 2;
                        } else if let Some(right) = &right
                            && right.valid[lane]
                        {
                            certificate.valid[lane] = true;
                            certificate.matched[lane] = right.matched[lane];
                            certificate.cost[lane] = left.cost[lane] + right.cost[lane] + 2;
                        }
                    }
                }
                Instruction::Builtin(b) => {
                    use grameo::Builtin;
                    let accepts: Option<fn(u8) -> bool> = match b {
                        Builtin::Ascii => Some(|b| b < 128),
                        Builtin::AsciiDigit => Some(|b| b.is_ascii_digit()),
                        Builtin::AsciiAlpha => Some(|b| b.is_ascii_alphabetic()),
                        Builtin::AsciiAlphaLower => Some(|b| b.is_ascii_lowercase()),
                        Builtin::AsciiAlphaUpper => Some(|b| b.is_ascii_uppercase()),
                        Builtin::AsciiAlphanumeric => Some(|b| b.is_ascii_alphanumeric()),
                        Builtin::AsciiHexDigit => Some(|b| b.is_ascii_hexdigit()),
                        Builtin::AsciiNonzeroDigit => Some(|b| (b'1'..=b'9').contains(&b)),
                        Builtin::AsciiBinDigit => Some(|b| matches!(b, b'0' | b'1')),
                        Builtin::AsciiOctDigit => Some(|b| (b'0'..=b'7').contains(&b)),
                        Builtin::XidStart if self.ascii_xid => Some(|b| b.is_ascii_alphabetic()),
                        Builtin::XidContinue if self.ascii_xid => Some(|b| b.is_ascii_alphanumeric() || b == b'_'),
                        _ => None,
                    };
                    if let Some(accepts) = accepts {
                        certificate.valid.fill(true);
                        certificate.cost.fill(1);
                        for byte in 0u8..=255 {
                            certificate.matched[usize::from(byte)] = accepts(byte);
                            if matches!(b,Builtin::XidStart|Builtin::XidContinue) && byte>=128 { certificate.valid[usize::from(byte)]=false; }
                        }
                    }
                }
                Instruction::Literal {
                    text,
                    insensitive: false,
                } if text.len() == 1 && text.is_ascii() => {
                    certificate.valid.fill(true);
                    certificate.cost.fill(1);
                    certificate.matched[usize::from(text.as_bytes()[0])] = true;
                }
                Instruction::Range { start, end } if start.is_ascii() && end.is_ascii() => {
                    certificate.valid.fill(true);
                    certificate.cost.fill(1);
                    for byte in 0u8..=255 {
                        certificate.matched[usize::from(byte)] =
                            byte >= start as u8 && byte <= end as u8;
                    }
                }
                _ => {}
            }
        }
        if !certificate.valid.iter().any(|&valid| valid) {
            self.bytes[id] = FragmentMemo::Unsupported;
            return None;
        }
        let certificate = std::rc::Rc::new(certificate);
        self.bytes[id] = FragmentMemo::Supported(std::rc::Rc::clone(&certificate));
        Some(certificate)
    }
}

type FragmentLiteralCertificate = (usize, Vec<(String, bool, usize)>);
fn fragment_literals(g: &CompiledGrammar, id: usize) -> Option<FragmentLiteralCertificate> {
    match g.instruction(id) {
        Instruction::Literal { text, insensitive } if !text.is_empty() && text.is_ascii() => {
            Some((1, vec![(text.into(), insensitive, 1)]))
        }
        Instruction::Group(child) => {
            let (failure, mut alternatives) = fragment_literals(g, child)?;
            for alternative in &mut alternatives {
                alternative.2 += 1;
            }
            Some((failure + 1, alternatives))
        }
        Instruction::Choice { left, right } => {
            let (left_failure, mut a) = fragment_literals(g, left)?;
            let (right_failure, b) = fragment_literals(g, right)?;
            for x in &mut a {
                x.2 += 2;
            }
            a.extend(
                b.into_iter()
                    .map(|(text, insensitive, cost)| (text, insensitive, cost + left_failure + 2)),
            );
            Some((left_failure + right_failure + 2, a))
        }
        _ => None,
    }
}
fn fragment_alternatives(alternatives: &[(String, bool, usize)]) -> String {
    alternatives.iter().map(|(text,insensitive,success_steps)|format!("::grameo::program::FragmentLiteral{{text:{text:?},insensitive:{insensitive},success_steps:{success_steps}}}")).collect::<Vec<_>>().join(",")
}
type FragmentGuardCertificate = (String, Vec<(String, bool, usize)>);
fn fragment_guard(
    g: &CompiledGrammar,
    id: usize,
    analysis: &mut FragmentAnalysis,
) -> Option<FragmentGuardCertificate> {
    let Instruction::Predicate {
        expression,
        positive: false,
    } = g.instruction(id)
    else {
        return None;
    };
    let (expression, call_groups) = fragment_unwrap(g, expression);
    let Instruction::Call(rule) = g.instruction(expression) else {
        return None;
    };
    if g.rule(rule).mode != grameo::RuleMode::Silent {
        return None;
    }
    let (root, groups) = fragment_unwrap(g, g.rule(rule).expression);
    let Instruction::Sequence { left, right } = g.instruction(root) else {
        return None;
    };
    let (failure, alternatives) = fragment_literals(g, left)?;
    if alternatives.len() > u16::MAX as usize {
        return None;
    }
    let candidates = (0u8..=255)
        .map(|byte| {
            let ids = alternatives
                .iter()
                .enumerate()
                .filter(|(_, (text, insensitive, _))| {
                    if *insensitive {
                        text.as_bytes()[0].eq_ignore_ascii_case(&byte)
                    } else {
                        text.as_bytes()[0] == byte
                    }
                })
                .map(|(index, _)| index.to_string())
                .collect::<Vec<_>>()
                .join(",");
            format!("&[{ids}]")
        })
        .collect::<Vec<_>>()
        .join(",");

    let (right, right_groups) = fragment_unwrap(g, right);
    let Instruction::Predicate {
        expression,
        positive: false,
    } = g.instruction(right)
    else {
        return None;
    };
    let boundary_certificate = analysis.certificate(g, expression)?;
    let mut boundary = Vec::new();
    for byte in 0..=255 {
        let (matched, cost, _) = boundary_certificate.probe(byte)?;
        boundary.push(if matched {
            u16::try_from(cost).ok()?
        } else {
            0
        });
    }
    Some((
        format!(
            "::grameo::program::FragmentGuard{{rule:{rule},alternatives:&[{}],candidates:&[{candidates}],failure_steps:{failure},boundary:&{boundary:?},choice_failure_fixed:{},boundary_fixed:{},boundary_failure_steps:{}}}",
            fragment_alternatives(&alternatives),
            8 + groups + call_groups,
            10 + groups + call_groups + right_groups,
            analysis.eof(g, expression)?.1
        ),
        alternatives,
    ))
}
fn literal_fragment_spec(
    g: &CompiledGrammar,
    rule: usize,
    analysis: &mut FragmentAnalysis,
) -> Option<String> {
    let (root, groups) = fragment_unwrap(g, g.rule(rule).expression);
    let Instruction::Sequence { left, right } = g.instruction(root) else {
        return None;
    };
    let (left, left_groups) = fragment_unwrap(g, left);
    let Instruction::Literal { text, insensitive } = g.instruction(left) else {
        return None;
    };
    if text.is_empty() || !text.is_ascii() {
        return None;
    }
    let (right, right_groups) = fragment_unwrap(g, right);
    let Instruction::Predicate {
        expression,
        positive: false,
    } = g.instruction(right)
    else {
        return None;
    };
    let boundary_certificate = analysis.certificate(g, expression)?;
    let mut boundary = Vec::new();
    for byte in 0..=255 {
        let (matched, cost, _) = boundary_certificate.probe(byte)?;
        boundary.push(if matched {
            u16::try_from(cost).ok()?
        } else {
            0
        });
    }
    let failure = analysis.eof(g, expression)?.1;
    let literal_failure = 4 + groups + left_groups;
    let fixed = 6 + groups + left_groups + right_groups;
    Some(format!(
        "::grameo::program::AtomicFragment{{literal:Some(&::grameo::program::AtomicLiteral{{text:{text:?},insensitive:{insensitive},boundary:&{boundary:?},literal_failure_steps:{literal_failure},boundary_fixed_steps:{fixed},boundary_failure_steps:{failure}}}),guard:None,suffix:None,first:&[0;256],tail:&[0;256],fixed_steps:0,failure_guard_fixed:0,failure_head_fixed:0,head_failure_steps:0,head_failure_safe:false,head_expected:&[::grameo::program::FragmentExpected::Literal({text:?})],expected:&[]}}"
    ))
}
fn fragment_spec(
    g: &CompiledGrammar,
    rule: usize,
    analysis: &mut FragmentAnalysis,
) -> Option<String> {
    if g.rule(rule).mode != grameo::RuleMode::Atomic {
        return None;
    }
    if let Some(spec) = literal_fragment_spec(g, rule, analysis) {
        return Some(spec);
    }
    fn flatten(g: &CompiledGrammar, id: usize, segments: &mut Vec<usize>) -> usize {
        match g.instruction(id) {
            Instruction::Group(child) => 1 + flatten(g, child, segments),
            Instruction::Sequence { left, right } => {
                3 + flatten(g, left, segments) + flatten(g, right, segments)
            }
            _ => {
                segments.push(id);
                0
            }
        }
    }
    let mut segments = Vec::new();
    let fixed = flatten(g, g.rule(rule).expression, &mut segments);
    let mut guard = "None".to_owned();
    let mut guard_origin = None;
    let mut guard_alternatives = Vec::new();
    if let Some(&first) = segments.first()
        && matches!(g.instruction(first), Instruction::Predicate { .. })
    {
        let (spec, alternatives) = fragment_guard(g, first, analysis)?;
        guard = format!("Some(&{spec})");
        guard_origin = Some(first);
        guard_alternatives = alternatives;
        segments.remove(0);
    }
    if !(2..=3).contains(&segments.len()) {
        return None;
    }
    let mut suffix = "None".to_owned();
    if segments.len() == 3 {
        let Instruction::Repeat {
            expression,
            min: 0,
            max: Some(1),
        } = g.instruction(segments[2])
        else {
            return None;
        };
        let (failure, alternatives) = fragment_literals(g, expression)?;
        if alternatives.iter().any(|(text, _, _)| text.len() != 1) {
            return None;
        }
        let expected = alternatives
            .iter()
            .map(|(text, _, _)| {
                format!("::grameo::program::FragmentExpected::Literal({text:?})")
            })
            .collect::<Vec<_>>()
            .join(",");
        suffix = format!(
            "Some(&::grameo::program::FragmentSuffix{{alternatives:&[{}],failure_steps:{failure},expected:&[{expected}]}})",
            fragment_alternatives(&alternatives)
        );
    }
    let Instruction::Repeat {
        expression,
        min: 0,
        max: None,
    } = g.instruction(segments[1])
    else {
        return None;
    };
    let (_, failed, expected) = analysis.eof(g, expression)?;
    fn prefix_fixed(g: &CompiledGrammar, id: usize, target: usize) -> Option<usize> {
        if id == target {
            return Some(0);
        }
        match g.instruction(id) {
            Instruction::Group(child) => prefix_fixed(g, child, target).map(|c| c + 1),
            Instruction::Sequence { left, right } => {
                if let Some(c) = prefix_fixed(g, left, target) {
                    Some(c + 3)
                } else {
                    prefix_fixed(g, right, target).map(|c| c + 3 + fixed_subtree(g, left))
                }
            }
            _ => None,
        }
    }
    fn fixed_subtree(g: &CompiledGrammar, id: usize) -> usize {
        match g.instruction(id) {
            Instruction::Group(child) => 1 + fixed_subtree(g, child),
            Instruction::Sequence { left, right } => {
                3 + fixed_subtree(g, left) + fixed_subtree(g, right)
            }
            _ => 0,
        }
    }
    let root = g.rule(rule).expression;
    let failure_guard_fixed = guard_origin
        .and_then(|id| prefix_fixed(g, root, id))
        .unwrap_or(0);
    let failure_head_fixed = prefix_fixed(g, root, segments[0])?;
    let (_, head_failure_steps, head_expected) = analysis.eof(g, segments[0])?;
    let first_certificate = analysis.certificate(g, segments[0])?;
    let tail_certificate = analysis.certificate(g, expression)?;
    let mut partial = false;
    let mut first = Vec::new();
    let mut tail = Vec::new();
    for byte in 0..=255 {
        let (matched, cost, _) = match first_certificate.probe(byte) { Some(v)=>v, None if analysis.ascii_xid && byte>=128=>{partial=true;first.push(u16::MAX);tail.push(u16::MAX);continue;}, None=>return None };
        first.push(if matched {
            u16::try_from(cost).ok()?
        } else {
            0
        });
        let (matched, cost, _) = match tail_certificate.probe(byte) { Some(v)=>v, None if analysis.ascii_xid && byte>=128=>{partial=true;tail.push(u16::MAX);continue;}, None=>return None };
        tail.push(if matched {
            u16::try_from(cost).ok()?
        } else {
            0
        });
    }
    if partial {
        // Only explicitly partial tables reserve the sentinel. Ordinary tables
        // retain every previously valid u16 cost, including u16::MAX.
        if first[..128].contains(&u16::MAX) || tail[..128].contains(&u16::MAX) { return None; }
        analysis.partial_rules.resize(g.syntax().rules.len(),false);
        analysis.partial_rules[rule]=true;
    }
    let head_failure_safe = guard_alternatives.iter().all(|(text, insensitive, _)| {
        let byte = text.as_bytes()[0];
        first[usize::from(byte)] != 0
            && (!insensitive
                || (first[usize::from(byte.to_ascii_lowercase())] != 0
                    && first[usize::from(byte.to_ascii_uppercase())] != 0))
    });
    Some(format!(
        "::grameo::program::AtomicFragment{{literal:None,guard:{guard},suffix:{suffix},first:&{first:?},tail:&{tail:?},fixed_steps:{},failure_guard_fixed:{failure_guard_fixed},failure_head_fixed:{failure_head_fixed},head_failure_steps:{head_failure_steps},head_failure_safe:{head_failure_safe},head_expected:&[{}],expected:&[{}]}}",
        fixed + 5 + failed,
        head_expected.join(","),
        expected.join(",")
    ))
}

fn write_instruction(out: &mut String, op: Instruction<'_>) -> std::fmt::Result {
    match op {
        Instruction::Literal { text, insensitive } => {
            write!(out, "Literal {{ text: {text:?}, insensitive: {insensitive} }}")
        }
        Instruction::Range { start, end } => write!(out, "Range {{ start: {start:?}, end: {end:?} }}"),
        Instruction::Call(id) => write!(out, "Call({id})"),
        Instruction::Builtin(builtin) => write!(out, "Builtin(::grameo::Builtin::{builtin:?})"),
        Instruction::Sequence { left, right } => {
            write!(out, "Sequence {{ left: {left}, right: {right} }}")
        }
        Instruction::Choice { left, right } => write!(out, "Choice {{ left: {left}, right: {right} }}"),
        Instruction::Repeat {
            expression,
            min,
            max,
        } => write!(out, "Repeat {{ expression: {expression}, min: {min}, max: {max:?} }}"),
        Instruction::Predicate {
            expression,
            positive,
        } => write!(out, "Predicate {{ expression: {expression}, positive: {positive} }}"),
        Instruction::Group(id) => write!(out, "Group({id})"),
        Instruction::Push(id) => write!(out, "Push({id})"),
        Instruction::PushLiteral(text) => write!(out, "PushLiteral({text:?})"),
        Instruction::PeekSlice { start, end } => {
            write!(out, "PeekSlice {{ start: {start}, end: {end:?} }}")
        }
    }
}

fn byte_class(class: Option<grameo::lexical::ByteClass>) -> String {
    class.map_or_else(
        || "None".into(),
        |class| format!("Some(::grameo::lexical::ByteClass({:?}))", class.0),
    )
}

#[cfg(test)]
mod fragment_byte_tests {
    use super::*;
    #[test]
    fn all_lanes_preserve_recursive_supported_costs() {
        let shapes = [
            r#"ASCII"#,
            r#"ASCII_DIGIT"#,
            r#"ASCII_ALPHA"#,
            r#"ASCII_ALPHA_LOWER"#,
            r#"ASCII_ALPHA_UPPER"#,
            r#"ASCII_ALPHANUMERIC"#,
            r#"ASCII_HEX_DIGIT"#,
            r#"ASCII_NONZERO_DIGIT"#,
            r#"ASCII_BIN_DIGIT"#,
            r#"ASCII_OCT_DIGIT"#,
            r#"ANY"#,
            r#""""#,
            r#"^"a""#,
            r#""ab""#,
            r#""é""#,
            r#"'a'..'z'"#,
            r#"(ASCII | ANY)"#,
            r#"(ASCII_DIGIT | ANY)"#,
            r#"(ANY | ASCII)"#,
            r#"(('a'..'z') | (ASCII_DIGIT | "_"))"#,
            r#"(("a") | (("b") | ANY))"#,
            r#"("a" | "b" | "c" | "d" | "e" | "f" | "g" | "h")"#,
            r#"((ASCII_DIGIT | ASCII_ALPHA) | (ASCII_ALPHA | "_"))"#,
        ];
        let mut comparisons = 0usize;
        for shape in shapes {
            let g = grameo::compile(&format!("entry = @{{ {shape} }}")).unwrap();
            let mut analysis = FragmentAnalysis::new(&g);
            for id in 0..g.syntax().expressions.len() {
                comparisons += 1;
                assert_eq!(
                    analysis.eof(&g, id),
                    fragment_class(&g, id, None),
                    "EOF {shape}, expr={id}"
                );
                for byte in 0u8..=255 {
                    comparisons += 1;
                    assert_eq!(
                        analysis.probe(&g, id, byte),
                        fragment_class(&g, id, Some(byte)),
                        "{shape}, expr={id}, byte={byte}"
                    );
                }
            }
        }
        println!("CLASS_PROBE_COMPARISONS\t{comparisons}");
    }
    #[test]
    fn supported_left_success_does_not_require_unsupported_right() {
        let g = grameo::compile("entry = @{ ASCII_DIGIT | ANY }").unwrap();
        let id = g.rule(0).expression;
        let mut analysis = FragmentAnalysis::new(&g);
        assert_eq!(analysis.probe(&g, id, b'1'), Some((true, 3, vec![])));
        assert_eq!(analysis.probe(&g, id, b'a'), None);
        let unsupported = (0..g.syntax().expressions.len())
            .find(|&id| {
                matches!(
                    g.instruction(id),
                    Instruction::Builtin(grameo::Builtin::Any)
                )
            })
            .unwrap();
        assert!(matches!(
            analysis.bytes[unsupported],
            FragmentMemo::Unsupported
        ));
        assert!(matches!(analysis.bytes[id], FragmentMemo::Supported(_)));
        assert!(analysis.certificate(&g, unsupported).is_none());
        assert!(matches!(
            analysis.bytes[unsupported],
            FragmentMemo::Unsupported
        ));
        assert_eq!(fragment_class(&g, id, None), None);
    }
    #[test]
    fn failed_choice_costs_remain_wide_until_output_conversion() {
        use grameo::grammar::{Expr, ExprKind};
        for depth in [14, 15] {
            let mut syntax = grameo::parse("entry = @{ ASCII_DIGIT }").unwrap();
            let mut previous = syntax.rules[0].expression;
            let span = syntax.expressions[previous].span;
            for _ in 0..depth {
                let parent = syntax.expressions.len();
                syntax.expressions.push(Expr {
                    kind: ExprKind::Choice {
                        left: previous,
                        right: previous,
                    },
                    span,
                });
                previous = parent;
            }
            let letter = syntax.expressions.len();
            syntax.expressions.push(Expr {
                kind: ExprKind::Literal {
                    text: "x".into(),
                    insensitive: false,
                },
                span,
            });
            let root = syntax.expressions.len();
            syntax.expressions.push(Expr {
                kind: ExprKind::Choice {
                    left: previous,
                    right: letter,
                },
                span,
            });
            syntax.rules[0].expression = root;
            // This unused parent prevents the existing root-only terminal-choice
            // accelerator from excluding the class being tested.
            syntax.expressions.push(Expr {
                kind: ExprKind::Group(root),
                span,
            });
            let g = grameo::resolve(syntax).unwrap();
            let mut analysis = FragmentAnalysis::new(&g);
            for byte in [b'0', b'x', 255] {
                let old = fragment_class(&g, root, Some(byte));
                let new = analysis.probe(&g, root, byte);
                assert_eq!(new, old);
                if byte == b'x' {
                    let cost = new.unwrap().1;
                    assert_eq!(u16::try_from(cost).is_ok(), depth == 14);
                }
            }
        }
    }
}

#[derive(Clone)]
struct NativeChoiceProof {
    bytes: [u64; 4],
    steps: usize,
    depth: usize,
    diagnostic_complete: bool,
    events: Vec<(usize, Option<usize>)>,
    calls: usize,
    contains_literal: bool,
}
fn native_choice_proof(
    grammar: &CompiledGrammar,
    id: usize,
    nesting: usize,
) -> Option<NativeChoiceProof> {
    if nesting >= 32 {
        return None;
    }
    match grammar.instruction(id) {
        Instruction::Literal { text, insensitive } => {
            let byte = *text.as_bytes().first()?;
            // ASCII prefixes are sufficient certificates even for UTF-8 input.
            // Empty/non-ASCII literals retain the ordinary matching path.
            if !byte.is_ascii() { return None; }
            let mut bytes = [0u64; 4];
            for candidate in [byte, if insensitive {byte.to_ascii_uppercase()} else {byte}, if insensitive {byte.to_ascii_lowercase()} else {byte}] {
                bytes[usize::from(candidate/64)] |= 1u64 << (candidate%64);
            }
            Some(NativeChoiceProof { bytes, steps: 1, depth: 0,
                diagnostic_complete: false, events: vec![(id,None)], calls: 0, contains_literal: true })
        }
        Instruction::Call(rule) => {
            let prefix = grammar.trivia_failure_prefix(rule)?;
            Some(NativeChoiceProof {
                bytes: prefix.bytes,
                steps: prefix.steps.checked_add(3)?,
                depth: prefix.depth.checked_add(1)?,
                diagnostic_complete: !grammar.failure_terminals(rule).is_empty(),
                events: vec![(id, Some(rule))],
                calls: 1,
                contains_literal: false,
            })
        }
        Instruction::Group(child) => {
            let mut proof = native_choice_proof(grammar, child, nesting + 1)?;
            if proof.contains_literal && grammar.terminal_choice(id).is_some() {return None;}
            proof.steps = proof.steps.checked_add(1)?;
            proof.events.insert(0, (id, None));
            Some(proof)
        }
        Instruction::Sequence { left, .. } => {
            // A certified failure in the left child never evaluates the right.
            // Preserve Eval, Sequence continuation and FinishExpr logical work.
            let mut proof = native_choice_proof(grammar, left, nesting + 1)?;
            proof.steps = proof.steps.checked_add(3)?;
            proof.events.insert(0, (id, None));
            Some(proof)
        }
        Instruction::Choice { left, right } => {
            let mut left = native_choice_proof(grammar, left, nesting + 1)?;
            let right = native_choice_proof(grammar, right, nesting + 1)?;
            if (left.contains_literal || right.contains_literal) && grammar.terminal_choice(id).is_some() {return None;}
            left.contains_literal |= right.contains_literal;
            if left.calls + right.calls > 32 {
                return None;
            }
            for index in 0..4 {
                left.bytes[index] |= right.bytes[index];
            }
            left.steps = left.steps.checked_add(right.steps)?.checked_add(2)?;
            left.depth = left.depth.max(right.depth);
            left.diagnostic_complete &= right.diagnostic_complete;
            left.calls += right.calls;
            left.events.insert(0, (id, None));
            left.events.extend(right.events);
            Some(left)
        }
        _ => None,
    }
}
fn emit_source_units(out: &mut String, grammar: &CompiledGrammar) {
    if grammar.code_space() == grameo::source_units::CodeSpace::Request {
        out.push_str("    fn code_space(&self) -> ::grameo::source_units::CodeSpace { ::grameo::source_units::CodeSpace::Request }\n");
    }
    let bindings: Vec<_> = grammar.source_unit_bindings().collect();
    // Preserve byte-identical default generated output when no hook is bound.
    if bindings.is_empty() { return; }
    out.push_str("    fn source_unit_binding(&self, rule:usize, selector:Option<usize>) -> Option<::grameo::source_units::Binding<'_>> {\n        use ::grameo::source_units::*;\n        match (rule, selector) {\n");
    for b in bindings {
        write!(out, "            ({}, {:?}) => Some(Binding {{rule:{},target:{},selector:{:?},loader:{:?},policy:Policy {{phase:Phase::{:?},repetition:Repetition::{:?},missing:Missing::{:?},scope:Scope::{:?}",
            b.rule, b.selector, b.rule, b.target, b.selector, b.loader,
            b.policy.phase, b.policy.repetition, b.policy.missing, b.policy.scope).unwrap();
        out.push_str("}}),\n");
    }
    out.push_str("            _ => None\n        }\n    }\n");
}

fn emit_native_choices(out: &mut String, grammar: &CompiledGrammar) {

    writeln!(out, "    #[inline]\n    fn native_hybrid_choice_prefix(&self, id: usize) -> Option<&::grameo::program::NativeChoicePrefix<'_>> {{ const PREFIXES: [Option<&::grameo::program::NativeChoicePrefix<'static>>; {}] = [", grammar.syntax().expressions.len()).unwrap();
    for id in 0..grammar.syntax().expressions.len() {
        let proof = match grammar.instruction(id) {
            Instruction::Choice { left, .. } => native_choice_proof(grammar, left, 0),
            _ => None,
        };
        let Some(proof) = proof else {
            out.push_str("        None,\n");
            continue;
        };
        write!(out, "        Some(&::grameo::program::NativeChoicePrefix {{ bytes: {:?}, steps: {}, depth: {}, diagnostic_complete: {}, events: &[", proof.bytes, proof.steps, proof.depth, proof.diagnostic_complete).unwrap();
        for (expression, rule) in proof.events {
            if let Some(rule) = rule {
                write!(out, "::grameo::program::NativeChoiceEvent::Call {{ expression: {expression}, rule: {rule} }},").unwrap();
            } else {
                write!(
                    out,
                    "::grameo::program::NativeChoiceEvent::Expression({expression}),"
                )
                .unwrap();
            }
        }
        out.push_str("] }),\n");
    }
    out.push_str("    ]; PREFIXES[id] }\n");
}

fn emit_native_rule_heads(out: &mut String, grammar: &CompiledGrammar) {
    writeln!(out, "    #[inline]\n    fn native_hybrid_rule_head(&self, rule: usize) -> Option<&::grameo::program::NativeRuleHead<'_>> {{ const HEADS: [Option<&::grameo::program::NativeRuleHead<'static>>; {}] = [",grammar.syntax().rules.len()).unwrap();
    for rule in 0..grammar.syntax().rules.len() {
        let mut id = grammar.rule(rule).expression;
        let mut sequences = Vec::new();
        while let Instruction::Sequence { left, right } = grammar.instruction(id) {
            sequences.push((id, right));
            id = left;
        }
        if sequences.len() <= 32
            && let Instruction::Call(callee) = grammar.instruction(id)
        {
            writeln!(out,"Some(&::grameo::program::NativeRuleHead {{ sequences: &{:?}, call_expression: {}, callee: {} }}),",sequences,id,callee).unwrap();
        } else {
            out.push_str("None,\n");
        }
    }
    out.push_str("    ]; HEADS[rule] }\n");
}

#[cfg(test)]mod native_literal_certification_tests {
 use super::*;
 fn proof(source:&str)->Option<NativeChoiceProof>{let grammar=grameo::compile(source).unwrap();native_choice_proof(&grammar,grammar.rule(0).expression,0)}
 #[test]fn ascii_literal_miss_certificate_keeps_case_and_logical_work(){
  let p=proof("entry = { ^\"Aλ\" }").unwrap();assert_eq!((p.steps,p.depth,p.calls,p.diagnostic_complete),(1,0,0,false));
  for byte in [b'A',b'a']{assert_ne!(p.bytes[usize::from(byte/64)]&(1u64<<(byte%64)),0);}
  assert_eq!(p.bytes[usize::from(b'b'/64)]&(1u64<<(b'b'%64)),0);
 }
 #[test]fn empty_and_unicode_first_literals_keep_fallback(){for source in["entry = { \"\" }","entry = { \"λ\" }"]{assert!(proof(source).is_none());}}
 #[test]fn terminal_shortcuts_keep_their_own_event_order(){assert!(proof("entry = { \"a\" | \"b\" | \"c\" | \"d\" | \"e\" | \"f\" | \"g\" | \"h\" }").is_none());}
 #[test]fn literal_prefixes_aggregate_across_nonterminal_choices(){let p=proof("entry = { (\"(\" ~ target) | target }\ntarget = { \"a\" }").unwrap();assert!(p.contains_literal);assert_eq!(p.calls,1);assert!(p.steps>4);}
 #[test]fn failed_left_literal_skips_only_certified_sequence_work(){let p=proof("entry = { \"(\" ~ target }\ntarget = { \"a\" }").unwrap();assert_eq!((p.steps,p.depth,p.calls),(4,0,0));assert_eq!(p.events.len(),2);assert!(p.contains_literal);}
}

#[cfg(test)]mod ascii_xid_fragment_tests {
 use super::*;
 #[test]fn unicode_lanes_require_whole_rule_fallback(){let g=grameo::compile("word = @{ XID_START ~ XID_CONTINUE* }").unwrap();let mut a=FragmentAnalysis::new(&g);assert!(fragment_spec(&g,0,&mut a).is_none());let mut a=FragmentAnalysis::new(&g);a.ascii_xid=true;let root=g.rule(0).expression;let Instruction::Sequence{left,right}=g.instruction(root)else{panic!()};assert!(a.certificate(&g,left).unwrap().probe(b'a').unwrap().0);assert!(!a.certificate(&g,left).unwrap().probe(b'_').unwrap().0);assert!(a.certificate(&g,left).unwrap().probe(0xc3).is_none());let Instruction::Repeat{expression,..}=g.instruction(right)else{panic!()};assert!(a.certificate(&g,expression).unwrap().probe(b'_').unwrap().0);assert!(a.certificate(&g,expression).unwrap().probe(0xce).is_none());}
 #[test]fn default_emission_keeps_unicode_identifiers_interpreted(){let g=grameo::compile("root = { word }\nword = @{ (XID_START | \"_\") ~ XID_CONTINUE* }").unwrap();let normal=emit_with_api(&g,ParserApi::HYBRID_CAPTURES).unwrap();let opted=grameo::compile_with_options("// @grameo optimize ascii_xid\nroot = { word }\nword = @{ (XID_START | \"_\") ~ XID_CONTINUE* }",grameo::CompileOptions{optimizations:true,..Default::default()}).unwrap();let native=emit_with_api(&opted,ParserApi::HYBRID_CAPTURES).unwrap();assert!(!normal.contains("FRAGMENT_RULE_1"));assert!(native.contains("FRAGMENT_RULE_1"));}
}
