//! Conservative progress/effect analysis. Uncertainty becomes a warning and a
//! runtime guard, never an optimistic purity claim or an unsound rejection.
use crate::grammar::{ExprId, RuleId};
use crate::{Builtin, CompiledGrammar, Diagnostic, ExprKind, Reference, RuleMode};

#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub struct Effects(u8);
impl Effects {
    pub const STACK_READ: Self = Self(1);
    pub const STACK_WRITE: Self = Self(2);
    pub const MODE_CHANGE: Self = Self(4);
    pub const CAPTURE: Self = Self(8);
    pub const LOOKAHEAD: Self = Self(16);
    pub fn contains(self, effect: Self) -> bool {
        self.0 & effect.0 == effect.0
    }
    pub fn union(self, other: Self) -> Self {
        Self(self.0 | other.0)
    }
    pub fn is_stack_dependent(self) -> bool {
        self.0 & 3 != 0
    }
}

#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub struct Facts {
    /// Overapproximation: true includes uncertain/input-dependent success.
    pub may_succeed_empty: bool,
    /// Proof of unconditional, zero-width success; false means unproven.
    pub always_empty: bool,
    pub effects: Effects,
}

#[derive(Debug, Clone, Default)]
pub struct Analysis {
    pub expressions: Vec<Facts>,
    pub rules: Vec<Facts>,
    pub errors: Vec<Diagnostic>,
    pub warnings: Vec<Diagnostic>,
}

pub fn analyze(grammar: &CompiledGrammar) -> Analysis {
    let syntax = grammar.syntax();
    let implicit_skip = grammar
        .scopes()
        .iter()
        .any(|scope| scope.whitespace.is_some() || scope.comment.is_some());
    let mut analysis = Analysis {
        expressions: vec![Facts::default(); syntax.expressions.len()],
        rules: vec![Facts::default(); syntax.rules.len()],
        ..Analysis::default()
    };
    // Topological arena passes, then rule-reference propagation to a fixed point.
    loop {
        let mut changed = false;
        for (id, expression) in syntax.expressions.iter().enumerate() {
            let get = |child: ExprId| analysis.expressions[child];
            let facts = match &expression.kind {
                ExprKind::Literal { text, .. } => Facts {
                    may_succeed_empty: text.is_empty(),
                    always_empty: text.is_empty(),
                    ..Facts::default()
                },
                ExprKind::Range { .. } => Facts::default(),
                ExprKind::Reference(_) => match grammar.reference(id).unwrap() {
                    Reference::Rule(rule) => analysis.rules[rule],
                    Reference::Builtin(builtin) => builtin_facts(builtin),
                },
                ExprKind::Sequence { left, right } => {
                    let left = get(*left);
                    let right = get(*right);
                    Facts {
                        may_succeed_empty: left.may_succeed_empty && right.may_succeed_empty,
                        always_empty: !implicit_skip && left.always_empty && right.always_empty,
                        effects: left.effects.union(right.effects),
                    }
                }
                ExprKind::Choice { left, right } => {
                    let left = get(*left);
                    let right = get(*right);
                    Facts {
                        may_succeed_empty: left.may_succeed_empty || right.may_succeed_empty,
                        always_empty: left.always_empty,
                        effects: left.effects.union(right.effects),
                    }
                }
                ExprKind::Repeat {
                    expression,
                    min,
                    max,
                } => {
                    let child = get(*expression);
                    Facts {
                        may_succeed_empty: *min == 0 || child.may_succeed_empty,
                        always_empty: *max == Some(0) || child.always_empty,
                        effects: if *max == Some(0) {
                            Effects::default()
                        } else {
                            child.effects
                        },
                    }
                }
                ExprKind::Predicate { expression, .. } => Facts {
                    may_succeed_empty: true,
                    always_empty: false,
                    effects: get(*expression).effects.union(Effects::LOOKAHEAD),
                },
                ExprKind::Group(child) => get(*child),
                ExprKind::Push(child) => {
                    let mut facts = get(*child);
                    facts.effects = facts.effects.union(Effects::STACK_WRITE);
                    facts
                }
                ExprKind::PushLiteral(_) => Facts {
                    may_succeed_empty: true,
                    always_empty: true,
                    effects: Effects::STACK_WRITE,
                },
                ExprKind::PeekSlice { .. } => Facts {
                    may_succeed_empty: true,
                    always_empty: false,
                    effects: Effects::STACK_READ,
                },
            };
            if facts != analysis.expressions[id] {
                analysis.expressions[id] = facts;
                changed = true;
            }
        }
        for (id, rule) in syntax.rules.iter().enumerate() {
            let mut facts = analysis.expressions[rule.expression];
            let scope = grammar.scopes()[grammar.rule_scopes[id]];
            let skip_effects = [scope.whitespace, scope.comment]
                .into_iter()
                .flatten()
                .fold(Effects::default(), |effects, rule| {
                    effects.union(analysis.rules[rule].effects)
                });
            if rule.mode != RuleMode::Silent {
                facts.effects = facts.effects.union(Effects::CAPTURE);
            }
            if matches!(
                rule.mode,
                RuleMode::Atomic | RuleMode::CompoundAtomic | RuleMode::NonAtomic
            ) {
                facts.effects = facts.effects.union(Effects::MODE_CHANGE);
            }
            if !matches!(rule.mode, RuleMode::Atomic | RuleMode::CompoundAtomic) {
                facts.effects = facts.effects.union(skip_effects);
            }
            if facts != analysis.rules[id] {
                analysis.rules[id] = facts;
                changed = true;
            }
        }
        if !changed {
            break;
        }
    }
    for expression in &syntax.expressions {
        if let ExprKind::Repeat {
            expression: child,
            max: None,
            ..
        } = expression.kind
        {
            let facts = analysis.expressions[child];
            if facts.always_empty {
                analysis.errors.push(Diagnostic::new(
                    "G012",
                    "unbounded repetition always succeeds without consuming input",
                    expression.span,
                ));
            } else if facts.may_succeed_empty {
                analysis.warnings.push(Diagnostic::new(
                    "W002",
                    "unbounded repetition may match empty input; runtime progress guard required",
                    expression.span,
                ));
            }
        }
    }
    let mut checked_skip = std::collections::HashSet::new();
    for scope in grammar.scopes() {
        for (name, rule) in [("WHITESPACE", scope.whitespace), ("COMMENT", scope.comment)] {
            if let Some(rule) = rule
                && checked_skip.insert(rule)
                && analysis.rules[rule].always_empty
            {
                analysis.errors.push(Diagnostic::new(
                    "G012",
                    format!("implicit {name} rule always succeeds without consuming input"),
                    syntax.rules[rule].span,
                ));
            }
        }
    }
    // Edges mark calls reachable before consuming input. Only calls proved to
    // be unconditionally attempted form the hard-error graph; others warn.
    let mut certain = vec![Vec::new(); syntax.rules.len()];
    let mut possible = vec![Vec::new(); syntax.rules.len()];
    for (rule, definition) in syntax.rules.iter().enumerate() {
        let mut pending = vec![(definition.expression, true)];
        while let Some((id, definite)) = pending.pop() {
            match syntax.expressions[id].kind {
                ExprKind::Reference(_) => {
                    if let Some(Reference::Rule(target)) = grammar.reference(id) {
                        possible[rule].push(target);
                        if definite && !analysis.rules[rule].effects.is_stack_dependent() {
                            certain[rule].push(target);
                        }
                    }
                }
                ExprKind::Sequence { left, right } => {
                    pending.push((left, definite));
                    if analysis.expressions[left].may_succeed_empty {
                        let no_skip = !implicit_skip
                            || matches!(
                                definition.mode,
                                RuleMode::Atomic | RuleMode::CompoundAtomic
                            );
                        pending.push((
                            right,
                            definite && no_skip && analysis.expressions[left].always_empty,
                        ));
                    }
                }
                ExprKind::Choice { left, right } => {
                    pending.push((left, definite));
                    pending.push((right, false));
                }
                ExprKind::Repeat {
                    expression, max, ..
                } => {
                    if max != Some(0) {
                        pending.push((expression, definite));
                    }
                }
                ExprKind::Group(child) | ExprKind::Push(child) => pending.push((child, definite)),
                ExprKind::Predicate { expression, .. } => pending.push((expression, definite)),
                _ => {}
            }
        }
        certain[rule].sort_unstable();
        certain[rule].dedup();
        possible[rule].sort_unstable();
        possible[rule].dedup();
    }
    for (graph, code, definite) in [(&certain, "G011", true), (&possible, "W001", false)] {
        for cycle in cycles(graph) {
            let names: Vec<_> = cycle
                .iter()
                .map(|id| syntax.rules[*id].name.as_str())
                .collect();
            let first = &syntax.rules[cycle[0]];
            let mut diagnostic = Diagnostic::new(
                code,
                format!(
                    "{}left recursion: {}",
                    if definite { "" } else { "possible " },
                    names.join(" -> ")
                ),
                first.name_span,
            );
            diagnostic.related.extend(cycle[1..].iter().map(|id| {
                (
                    syntax.rules[*id].name_span,
                    "rule in recursion cycle".into(),
                )
            }));
            if definite {
                analysis.errors.push(diagnostic);
            } else {
                analysis.warnings.push(diagnostic);
            }
        }
    }
    analysis
}

fn builtin_facts(builtin: Builtin) -> Facts {
    use Builtin::*;
    match builtin {
        Soi => Facts {
            may_succeed_empty: true,
            ..Facts::default()
        },
        Eoi => Facts {
            may_succeed_empty: true,
            effects: Effects::CAPTURE,
            ..Facts::default()
        },
        Peek | PeekAll => Facts {
            may_succeed_empty: true,
            effects: Effects::STACK_READ,
            ..Facts::default()
        },
        Pop | PopAll | Drop => Facts {
            may_succeed_empty: true,
            effects: Effects::STACK_READ.union(Effects::STACK_WRITE),
            ..Facts::default()
        },
        _ => Facts::default(),
    }
}

/// Iterative DFS: grammar recursion must not recurse on the Rust stack either.
fn cycles(graph: &[Vec<RuleId>]) -> Vec<Vec<RuleId>> {
    let mut colors = vec![0u8; graph.len()];
    let mut result = Vec::new();
    for start in 0..graph.len() {
        if colors[start] != 0 {
            continue;
        }
        let mut stack = vec![(start, 0usize)];
        colors[start] = 1;
        while let Some(&(node, next)) = stack.last() {
            if next == graph[node].len() {
                colors[node] = 2;
                stack.pop();
                continue;
            }
            stack.last_mut().unwrap().1 += 1;
            let target = graph[node][next];
            match colors[target] {
                0 => {
                    colors[target] = 1;
                    stack.push((target, 0));
                }
                1 => {
                    let index = stack.iter().position(|(node, _)| *node == target).unwrap();
                    let mut cycle: Vec<_> = stack[index..].iter().map(|(node, _)| *node).collect();
                    cycle.push(target);
                    result.push(cycle);
                }
                _ => {}
            }
        }
    }
    result
}
