//! Immutable admission for lexical shortcuts used with preparation hooks.
use super::Binding;
use crate::{grammar::{ExprId,RuleId,RuleMode},program::{Instruction,Program}};
use std::collections::HashSet;

fn children(instruction:Instruction<'_>,expressions:&mut Vec<ExprId>,rules:&mut Vec<RuleId>) {
    match instruction {
        Instruction::Call(rule)=>rules.push(rule),
        Instruction::Sequence {left,right}|Instruction::Choice {left,right}=>expressions.extend([left,right]),
        Instruction::Repeat {expression,..}|Instruction::Predicate {expression,..}|
        Instruction::Group(expression)|Instruction::Push(expression)=>expressions.push(expression),
        _=>{},
    }
}

/// Certify atomic rules whose own callback remains visible and whose children
/// cannot reach a request root, selector or operand. Predicates are included
/// conservatively. This inspects grammar instructions, never source or AST.
/// Cache the result only for the same immutable Program and binding set.
pub fn opaque_rules_preserving_requests(program:&impl Program,entries:&[RuleId],bindings:&[Binding<'_>])->Vec<RuleId> {
    let observed:HashSet<_>=bindings.iter().flat_map(|b|[Some(b.rule),Some(b.target),b.selector]).flatten().collect();
    let mut rules=entries.to_vec();let mut reachable=HashSet::new();let mut seen=HashSet::new();let mut expressions=Vec::new();
    while let Some(rule)=rules.pop() {
        if !reachable.insert(rule) {continue;}
        expressions.push(program.rule(rule).expression);
        while let Some(expression)=expressions.pop() {
            if seen.insert(expression) {children(program.instruction(expression),&mut expressions,&mut rules);}
        }
    }
    let mut admitted=Vec::new();
    for rule in reachable {
        let definition=program.rule(rule);
        if definition.mode!=RuleMode::Atomic || !program.atomic_expression(definition.expression) {continue;}
        expressions.push(definition.expression);rules.clear();seen.clear();let mut safe=true;
        loop {
            while let Some(expression)=expressions.pop() {
                if seen.insert(expression) {children(program.instruction(expression),&mut expressions,&mut rules);}
            }
            let Some(child)=rules.pop() else {break;};
            if observed.contains(&child) {safe=false;break;}
            expressions.push(program.rule(child).expression);
        }
        expressions.clear();
        if safe {admitted.push(rule);}
    }
    admitted.sort_unstable();admitted
}
