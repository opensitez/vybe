use crate::RuleMode;
use crate::program::{Instruction, Program};
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct EmptySuffix {
    pub owner: usize,
    pub owner_body: usize,
    pub repetition: usize,
    pub child: usize,
    pub scope: usize,
    pub nullable_heads: [u64; 4],
    pub mandatory_heads: [u64; 4],
    pub calls: Vec<usize>,
}
fn bit(mask: &mut [u64; 4], b: u8) {
    mask[(b / 64) as usize] |= 1u64 << (b % 64);
}
fn literals<P: Program + ?Sized>(p: &P, id: usize, mask: &mut [u64; 4]) -> Option<()> {
    let mut stack = vec![id];
    let mut steps = 0;
    while let Some(id) = stack.pop() {
        steps += 1;
        if steps > 4096 {
            return None;
        }
        match p.instruction(id) {
            Instruction::Group(x) => stack.push(x),
            Instruction::Choice { left, right } => {
                stack.push(right);
                stack.push(left);
            }
            Instruction::Literal { text, insensitive } => {
                let b = *text.as_bytes().first()?;
                bit(mask, b);
                if insensitive && b.is_ascii_alphabetic() {
                    bit(mask, b.to_ascii_uppercase());
                    bit(mask, b.to_ascii_lowercase());
                }
            }
            _ => return None,
        }
    }
    Some(())
}
fn flatten<P: Program + ?Sized>(p: &P, id: usize) -> Option<Vec<usize>> {
    let mut stack = vec![id];
    let mut out = vec![];
    let mut steps = 0;
    while let Some(id) = stack.pop() {
        steps += 1;
        if steps > 4096 {
            return None;
        }
        match p.instruction(id) {
            Instruction::Group(x) => stack.push(x),
            Instruction::Sequence { left, right } => {
                stack.push(right);
                stack.push(left);
            }
            _ => out.push(id),
        }
    }
    Some(out)
}
fn ungroup<P: Program + ?Sized>(p: &P, mut id: usize) -> Option<usize> {
    for _ in 0..64 {
        match p.instruction(id) {
            Instruction::Group(x) => id = x,
            _ => return Some(id),
        }
    }
    None
}
fn proof<P: Program + ?Sized>(p: &P, owner: usize) -> Option<EmptySuffix> {
    let r = p.rule(owner);
    if p.source_pratt(owner).is_some() {
        return None;
    }
    let mut stack = vec![r.expression];
    let mut found = None;
    let mut steps = 0;
    while let Some(id) = stack.pop() {
        steps += 1;
        if steps > 4096 {
            return None;
        }
        match p.instruction(id) {
            Instruction::Group(x) => stack.push(x),
            Instruction::Sequence { left, right } | Instruction::Choice { left, right } => {
                stack.push(left);
                stack.push(right);
            }
            Instruction::Repeat {
                min: 0,
                max: Some(1),
                ..
            } => {
                if found.replace(id).is_some() {
                    return None;
                }
            }
            _ => {}
        }
    }
    let repetition = found?;
    let Instruction::Repeat {
        expression: child, ..
    } = p.instruction(repetition)
    else {
        unreachable!()
    };
    let segments = flatten(p, child)?;
    let mut nullable_heads = [0; 4];
    let mut mandatory_heads = [0; 4];
    let mut calls = vec![];
    let mut mandatory = false;
    for id in segments {
        let id = ungroup(p, id)?;
        let (body, called) = match p.instruction(id) {
            Instruction::Call(rule) => {
                let rr = p.rule(rule);
                if rr.scope != r.scope || p.source_pratt(rule).is_some() {
                    return None;
                }
                (ungroup(p, rr.expression)?, Some(rule))
            }
            _ => (id, None),
        };
        if let Instruction::Repeat {
            expression,
            min: 0,
            max,
        } = p.instruction(body)
        {
            if max == Some(0) {
                return None;
            }
            if let Some(rule) = called {
                if p.rule(rule).mode != RuleMode::Silent {
                    return None;
                }
                calls.push(rule);
            }
            literals(p, expression, &mut nullable_heads)?;
        } else {
            literals(p, body, &mut mandatory_heads)?;
            if let Some(rule) = called {
                calls.push(rule);
            }
            mandatory = true;
            break;
        }
    }
    if !mandatory {
        return None;
    }
    Some(EmptySuffix {
        owner,
        owner_body: r.expression,
        repetition,
        child,
        scope: r.scope,
        nullable_heads,
        mandatory_heads,
        calls,
    })
}

fn required_heads<P: Program + ?Sized>(
    p: &P,
    id: usize,
    mask: &mut [u64; 4],
    depth: usize,
) -> Option<()> {
    if depth >= 64 {
        return None;
    }
    match p.instruction(id) {
        Instruction::Literal { text, insensitive } => {
            let b = *text.as_bytes().first()?;
            bit(mask, b);
            if insensitive && b.is_ascii_alphabetic() {
                bit(mask, b.to_ascii_uppercase());
                bit(mask, b.to_ascii_lowercase());
            }
            Some(())
        }
        Instruction::Group(x) => required_heads(p, x, mask, depth + 1),
        Instruction::Choice { left, right } => {
            required_heads(p, left, mask, depth + 1)?;
            required_heads(p, right, mask, depth + 1)
        }
        Instruction::Sequence { left, .. } => required_heads(p, left, mask, depth + 1),
        Instruction::Repeat {
            expression, min, ..
        } if min > 0 => required_heads(p, expression, mask, depth + 1),
        Instruction::Call(rule) => {
            if p.source_pratt(rule).is_some() {
                return None;
            }
            required_heads(p, p.rule(rule).expression, mask, depth + 1)
        }
        _ => None,
    }
}
pub fn scope_heads<P: Program + ?Sized>(p: &P, scope: usize) -> Option<[u64; 4]> {
    let s = p.scope(scope);
    let mut mask = [0; 4];
    for rule in [s.whitespace, s.comment].into_iter().flatten() {
        required_heads(p, p.rule(rule).expression, &mut mask, 0)?;
    }
    Some(mask)
}

pub(crate) fn certify(
    g: &crate::CompiledGrammar,
    owner: usize,
) -> Option<crate::program::EmptySuffixHead> {
    let p = proof(g, owner)?;
    let mut first = p.nullable_heads;
    for (a, b) in first.iter_mut().zip(p.mandatory_heads) {
        *a |= b;
    }
    for (a, b) in first.iter_mut().zip(scope_heads(g, p.scope)?) {
        *a |= b;
    }
    if p.calls.len() > 8 {
        return None;
    }
    let mut calls = [0; 8];
    calls[..p.calls.len()].copy_from_slice(&p.calls);
    Some(crate::program::EmptySuffixHead {
        owner: p.owner,
        owner_body: p.owner_body,
        repetition: p.repetition,
        child: p.child,
        scope: p.scope,
        first,
        calls,
        calls_len: p.calls.len() as u8,
        depth: 66,
    })
}
