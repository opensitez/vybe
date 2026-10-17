//! Conservative opt-in compile-time zero-consumption failure certificates.
use grameo::program::{Instruction, Program};
use grameo::{Builtin, CompiledGrammar, RuleMode};
#[derive(Clone, Copy)]
struct Env {
    scope: usize,
    mode: u8,
    skipping: bool,
}
#[derive(Clone, Debug)]
struct Proof {
    zero_success: bool,
    steps: usize,
    depth: usize,
    expected: Vec<usize>,
    rules: Vec<usize>,
}
impl Proof {
    fn terminal(zero_success: bool, expected: Option<usize>) -> Self {
        Self {
            zero_success,
            steps: 1,
            depth: 0,
            expected: expected.into_iter().collect(),
            rules: vec![],
        }
    }
    fn append(&mut self, other: Self) {
        self.steps += other.steps;
        self.depth = self.depth.max(other.depth);
        self.expected.extend(other.expected);
        self.rules.extend(other.rules);
    }
}
struct Prover<'a> {
    g: &'a CompiledGrammar,
    byte: Option<u8>,
    active: Vec<usize>,
    fuel: usize,
    expression_depth: usize,
}
impl Prover<'_> {
    fn rule(&mut self, id: usize, incoming: Env) -> Option<Proof> {
        if self.active.contains(&id) || self.active.len() >= 128 {
            return None;
        }
        self.active.push(id);
        let r = self.g.rule(id);
        let scope = self.g.scope(r.scope);
        let mut mode = match r.mode {
            RuleMode::Atomic => 1,
            RuleMode::CompoundAtomic => 2,
            RuleMode::NonAtomic => 0,
            _ => incoming.mode,
        };
        if scope.whitespace == Some(id) || scope.comment == Some(id) {
            mode = if r.mode == RuleMode::CompoundAtomic {
                2
            } else {
                1
            };
        }
        let body = self.expr(
            r.expression,
            Env {
                scope: r.scope,
                mode,
                skipping: incoming.skipping,
            },
        );
        self.active.pop();
        let mut p = body?;
        // A successful named zero-width capture is conservatively unsupported.
        if p.zero_success && r.mode != RuleMode::Silent {
            return None;
        }
        p.steps += 2;
        p.depth += 1;
        p.rules.push(id);
        Some(p)
    }
    fn skip(&mut self, env: Env) -> Option<Proof> {
        let scope = self.g.scope(env.scope);
        let mut p = Proof::terminal(true, None); // Skip task
        let prefix = scope.whitespace_prefix_class;
        if prefix.is_some_and(|class| self.byte.is_some_and(|b| class.contains(b))) {
            return None;
        }
        // The zero-byte prefix scanner itself checks rule depth.
        if prefix.is_some() {
            p.depth = 1;
        }
        let whitespace_done = prefix.is_some() && scope.whitespace_class.is_some();
        for (id, done) in [(scope.whitespace, whitespace_done), (scope.comment, false)] {
            if done {
                continue;
            }
            p.steps += 1; // AfterWhitespace / AfterComment
            if let Some(id) = id {
                let x = self.rule(
                    id,
                    Env {
                        skipping: true,
                        ..env
                    },
                )?;
                if x.zero_success {
                    return None;
                } // NonProgress or successful trivia is not a zero miss.
                p.append(x);
            }
        }
        p.zero_success = true;
        Some(p)
    }
    fn expr(&mut self, id: usize, env: Env) -> Option<Proof> {
        if self.fuel == 0 || self.expression_depth >= 128 {
            return None;
        }
        self.fuel -= 1;
        self.expression_depth += 1;
        let result = self.expr_inner(id, env);
        self.expression_depth -= 1;
        result
    }
    fn expr_inner(&mut self, id: usize, env: Env) -> Option<Proof> {
        match self.g.instruction(id) {
            Instruction::Literal { text, insensitive } => {
                if text.is_empty() {
                    return Some(Proof::terminal(true, None));
                }
                let first = *text.as_bytes().first()?;
                if self
                    .byte
                    .is_some_and(|b| b == first || insensitive && b.eq_ignore_ascii_case(&first))
                {
                    return None;
                }
                Some(Proof::terminal(false, (!env.skipping).then_some(id)))
            }
            Instruction::Range { start, end } => {
                if self
                    .byte
                    .is_some_and(|b| (start..=end).contains(&char::from(b)))
                {
                    return None;
                }
                Some(Proof::terminal(false, (!env.skipping).then_some(id)))
            }
            Instruction::Builtin(b) => {
                let accepts = match b {
                    Builtin::Ascii => self.byte.is_some(),
                    Builtin::AsciiDigit => self.byte.is_some_and(|x| x.is_ascii_digit()),
                    Builtin::AsciiAlpha => self.byte.is_some_and(|x| x.is_ascii_alphabetic()),
                    Builtin::AsciiAlphanumeric => {
                        self.byte.is_some_and(|x| x.is_ascii_alphanumeric())
                    }
                    Builtin::Newline => self.byte.is_some_and(|x| x == b'\n' || x == b'\r'),
                    _ => return None,
                };
                if accepts {
                    return None;
                }
                Some(Proof::terminal(false, (!env.skipping).then_some(id)))
            }
            Instruction::Call(rule) => {
                let mut p = self.rule(rule, env)?;
                p.steps += 1;
                Some(p)
            }
            Instruction::Group(child) => {
                let mut p = self.expr(child, env)?;
                p.steps += 1;
                Some(p)
            }
            Instruction::Sequence { left, right } => {
                let mut p = self.expr(left, env)?;
                p.steps += 3; // Eval Sequence, Sequence continuation, FinishExpr
                if !p.zero_success {
                    return Some(p);
                }
                if env.mode == 0 && !env.skipping {
                    p.append(self.skip(env)?);
                    p.steps += 1;
                } // AfterSkip
                let right = self.expr(right, env)?;
                p.zero_success = right.zero_success;
                p.append(right);
                Some(p)
            }
            Instruction::Choice { left, right } => {
                let mut p = self.expr(left, env)?;
                p.steps += 2; // Eval Choice + Choice continuation
                if p.zero_success {
                    return Some(p);
                }
                let right = self.expr(right, env)?;
                p.zero_success = right.zero_success;
                p.append(right);
                Some(p)
            }
            Instruction::Repeat {
                expression,
                min,
                max,
            } => {
                if max == Some(0) {
                    return Some(Proof {
                        steps: 3,
                        ..Proof::terminal(true, None)
                    });
                }
                let mut p = self.expr(expression, env)?;
                if p.zero_success {
                    return None;
                }
                // A failed first attempt finishes an unbounded repeat or a required repeat.
                // Optional {0,1} also executes its final bounded Repeat continuation.
                if min == 0 && max.is_some_and(|n| n != 1) {
                    return None;
                }
                p.steps += 5 + usize::from(min == 0 && max == Some(1));
                p.zero_success = min == 0;
                Some(p)
            }
            // Stackful, predicates, consuming prefixes and unsupported bodies fall back.
            _ => None,
        }
    }
}
fn nullable_head(g: &CompiledGrammar, id: usize, depth: usize) -> bool {
    if depth >= 128 {
        return false;
    }
    match g.instruction(id) {
        Instruction::Repeat { min: 0, .. } => true,
        Instruction::Literal { text: "", .. } => true,
        Instruction::Sequence { left, .. } | Instruction::Group(left) => {
            nullable_head(g, left, depth + 1)
        }
        Instruction::Call(rule) => nullable_head(g, g.rule(rule).expression, depth + 1),
        Instruction::Choice { left, right } => {
            nullable_head(g, left, depth + 1) || nullable_head(g, right, depth + 1)
        }
        _ => false,
    }
}
fn terminal(g: &CompiledGrammar, id: usize) -> String {
    match g.instruction(id) {
        Instruction::Literal { text, .. } => {
            format!("::grameo::program::FragmentExpected::Literal({text:?})")
        }
        Instruction::Range { start, end } => {
            format!("::grameo::program::FragmentExpected::Range({start:?},{end:?})")
        }
        Instruction::Builtin(b) => {
            format!("::grameo::program::FragmentExpected::Builtin(::grameo::Builtin::{b:?})")
        }
        _ => unreachable!("certified terminal"),
    }
}
pub(super) fn emit(out: &mut String, g: &CompiledGrammar) {
    use std::fmt::Write;
    if !g.optimization_selection().nullable_failure_heads {
        return;
    }
    let mut tables = Vec::new();
    for rule in 0..g.syntax().rules.len() {
        if !nullable_head(g, g.rule(rule).expression, 0) {
            continue;
        }
        let mut unique: Vec<(Proof, Vec<String>)> = Vec::new();
        let mut indices = [[0u16; 129]; 3];
        for mode in 0..3 {
            for (byte_index, byte) in (0u8..128).map(Some).chain([None]).enumerate() {
                let mut prover = Prover {
                    g,
                    byte,
                    active: vec![],
                    fuel: 32_768,
                    expression_depth: 0,
                };
                let Some(mut p) = prover
                    .rule(
                        rule,
                        Env {
                            scope: g.rule(rule).scope,
                            mode: mode as u8,
                            skipping: false,
                        },
                    )
                    .filter(|p| !p.zero_success)
                else {
                    continue;
                };
                let mut terms = Vec::new();
                for &id in &p.expected {
                    let term = terminal(g, id);
                    if !terms.contains(&term) {
                        terms.push(term);
                    }
                }
                let mut seen = std::collections::HashSet::new();
                p.rules.retain(|r| seen.insert(*r));
                let index = unique
                    .iter()
                    .position(|(old, t)| {
                        old.steps == p.steps
                            && old.depth == p.depth
                            && *t == terms
                            && old.rules == p.rules
                    })
                    .unwrap_or_else(|| {
                        unique.push((p, terms));
                        unique.len() - 1
                    });
                indices[mode][byte_index] =
                    u16::try_from(index + 1).expect("at most 387 ASCII/mode proofs");
            }
        }
        if !unique.is_empty() {
            tables.push((rule, unique, indices));
        }
    }
    if tables.is_empty() {
        return;
    }
    out.push_str("    fn match_observers_are_noop(&self) -> bool { true }\n");
    out.push_str("    fn nullable_failure_prefix(&self, rule:usize, mode:u8, first:Option<u8>) -> Option<&::grameo::program::NullableFailurePrefix<'_>> {\n        if mode>2 {return None;}\n        let byte=match first {Some(b) if b<128=>usize::from(b),None=>128,_=>return None};\n        match rule {\n");
    for (rule, unique, indices) in tables {
        writeln!(out, "            {rule} => {{").unwrap();
        for (index, (p, terms)) in unique.iter().enumerate() {
            writeln!(out,"                static P{index}: ::grameo::program::NullableFailurePrefix<'static> = ::grameo::program::NullableFailurePrefix {{steps:{},depth:{},expected:&[{}],unobserved_rules:&{:?}}};",p.steps,p.depth,terms.join(","),p.rules).unwrap();
        }
        writeln!(out,"                const MAP:[[u16;129];3]={indices:?};\n                match MAP[usize::from(mode)][byte] {{").unwrap();
        for index in 0..unique.len() {
            writeln!(out, "                    {}=>Some(&P{index}),", index + 1).unwrap();
        }
        out.push_str("                    _=>None\n                }\n            },\n");
    }
    out.push_str("            _=>None\n        }\n    }\n");
}

#[cfg(test)]
mod tests {
    use super::*;
    use grameo::{MatchOptions, ParseErrorKind};
    fn check_case(g: &CompiledGrammar, entry: &str, path: &str, checks: &mut usize) {
        let target = g.rule_id(entry).unwrap();
        let mut covered = 0;
        let mut costs = std::collections::BTreeSet::new();
        for byte in (0u8..128).map(Some).chain([None]) {
            let mut prover = Prover {
                g,
                byte,
                active: vec![],
                fuel: 32_768,
                expression_depth: 0,
            };
            let Some(proof) = prover.rule(
                target,
                Env {
                    scope: 0,
                    mode: 0,
                    skipping: false,
                },
            ) else {
                continue;
            };
            if proof.zero_success {
                continue;
            }
            let source = byte.map(|b| char::from(b).to_string()).unwrap_or_default();
            let actual = grameo::engine::parse_program(g, entry, &source, MatchOptions::default())
                .unwrap_err();
            assert_eq!(actual.kind, ParseErrorKind::Syntax);
            assert_eq!(actual.offset, 0, "{path} {byte:?}");
            let mut seen = std::collections::HashSet::new();
            let expected: Vec<_> = proof
                .expected
                .iter()
                .map(|&id| match g.instruction(id) {
                    Instruction::Literal { text, .. } => format!("{text:?}"),
                    Instruction::Range { start, end } => format!("{start:?}..{end:?}"),
                    Instruction::Builtin(Builtin::Newline) => "NEWLINE".into(),
                    Instruction::Builtin(b) => format!("{b:?}"),
                    _ => unreachable!(),
                })
                .filter(|s| seen.insert(s.clone()))
                .collect();
            assert_eq!(actual.expected, expected, "expectation {path} {byte:?}");
            for budget in 0..=proof.steps + 1 {
                let error = grameo::engine::parse_program(
                    g,
                    entry,
                    &source,
                    MatchOptions {
                        max_steps: budget,
                        max_rule_depth: 4096,
                    },
                )
                .unwrap_err();
                let expected = if budget < proof.steps {
                    ParseErrorKind::WorkLimit
                } else {
                    ParseErrorKind::Syntax
                };
                assert_eq!(
                    error.kind, expected,
                    "budget {path} byte={byte:?} budget={budget} proof={proof:?}"
                );
                *checks += 1;
            }
            for depth in 0..=proof.depth + 1 {
                let error = grameo::engine::parse_program(
                    g,
                    entry,
                    &source,
                    MatchOptions {
                        max_steps: usize::MAX,
                        max_rule_depth: depth,
                    },
                )
                .unwrap_err();
                let expected = if depth < proof.depth {
                    ParseErrorKind::DepthLimit
                } else {
                    ParseErrorKind::Syntax
                };
                assert_eq!(
                    error.kind, expected,
                    "depth {path} byte={byte:?} depth={depth} proof={proof:?}"
                );
                *checks += 1;
            }
            costs.insert((proof.steps, proof.depth));
            covered += 1;
        }
        println!(
            "PROOF file={path} certified_ascii_eof_cells={covered}/129 checks={checks} cost_depth={costs:?}"
        );
    }

    #[test]
    fn exact_work_depth_expectations_and_scopes() {
        let mut checks = 0;
        for (name, source) in [
            (
                "simple",
                r###"WHITESPACE = _{ " " | "\t" }
COMMENT = _{ "//" ~ (!"\n" ~ ANY)* }
NEWLINE = { "\n" | "\r\n" }
lead = _{ NEWLINE* }
root = { lead ~ ("?." | ".") | "(" | "[" | "`" }
"###,
            ),
            (
                "atomic",
                r###"NEWLINE = { "\n" | "\r\n" } lead = _{ NEWLINE* } root = @{ lead ~ ("?." | ".") | "(" }
"###,
            ),
            (
                "compound",
                r###"WHITESPACE = _{ " " } NEWLINE = { "\n" } lead = _{ NEWLINE* } root = ${ lead ~ "." | "(" }
"###,
            ),
            (
                "optional",
                r###"WHITESPACE = _{ " " } root = { "x"? ~ "." | "(" }
"###,
            ),
            (
                "empty",
                r###"root = { "" ~ "." | "(" }
"###,
            ),
            (
                "notrivia",
                r###"NEWLINE = { "\n" } lead = _{ NEWLINE* } root = { lead ~ "." | "(" }
"###,
            ),
            (
                "unknown",
                r###"lead = _{ PUSH("x"?) } root = { lead ~ "." }
"###,
            ),
        ] {
            let g = grameo::compile(source).unwrap();
            check_case(&g, "root", name, &mut checks);
        }
        for (label, outer, inner) in [
            (
                "scoped_nonatomic",
                r##"WHITESPACE = _{ "_" } root = @{ imported }"##,
                r##"WHITESPACE = _{ " " } COMMENT = _{ "#" ~ ANY* } NEWLINE = { "\n" } lead = _{ NEWLINE* } word = !{ lead ~ "." | "(" }"##,
            ),
            (
                "scoped_atomic",
                r##"WHITESPACE = _{ " " } root = { imported }"##,
                r##"WHITESPACE = _{ "_" } NEWLINE = { "\n" } lead = _{ NEWLINE* } word = @{ lead ~ "." | "(" }"##,
            ),
            (
                "scoped_compound",
                r##"WHITESPACE = _{ " " } root = { imported }"##,
                r##"WHITESPACE = _{ "_" } NEWLINE = { "\n" } lead = _{ NEWLINE* } word = ${ lead ~ "." | "(" }"##,
            ),
            (
                "scoped_prefix_fallback",
                r##"WHITESPACE = _{ " " } root = { imported }"##,
                r##"WHITESPACE = _{ "_" | "xx" } NEWLINE = { "\n" } lead = _{ NEWLINE* } word = { lead ~ "." | "(" }"##,
            ),
        ] {
            use grameo::modules::{Import, Module};
            let linked = grameo::modules::compile(&[
                Module {
                    name: "outer",
                    source: outer,
                    exports: &["root"],
                    imports: &[Import {
                        local: "imported",
                        module: "inner",
                        rule: "word",
                    }],
                },
                Module {
                    name: "inner",
                    source: inner,
                    exports: &["word"],
                    imports: &[],
                },
            ])
            .unwrap();
            check_case(linked.grammar(), "outer::root", label, &mut checks);
        }
        println!("generic certificate oracle checks={checks}");
    }
    #[test]
    fn directive_is_explicit_and_default_generation_unchanged() {
        let source = r#"root = { lead ~ "x" } lead = _{ NEWLINE* }"#;
        let annotated = format!("// @grameo optimize nullable_failure_heads\n{source}");
        let default = grameo::compile(&source).unwrap();
        let ignored = grameo::compile(&annotated).unwrap();
        assert!(!ignored.optimization_selection().nullable_failure_heads);
        assert_eq!(
            crate::emit(&default).unwrap(),
            crate::emit(&ignored).unwrap()
        );
        let opted = grameo::compile_with_options(
            &annotated,
            grameo::CompileOptions {
                optimizations: true,
                ..Default::default()
            },
        )
        .unwrap();
        assert!(opted.optimization_selection().nullable_failure_heads);
        assert!(
            crate::emit(&opted)
                .unwrap()
                .contains("fn nullable_failure_prefix")
        );
    }
}
