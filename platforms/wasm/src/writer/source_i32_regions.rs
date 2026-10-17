//! Physical i32 results consumed inside an uninterrupted integer expression.
//! Local reads stay boxed until their operator runs, preserving trap order.
use super::*;
use std::collections::BTreeSet;

#[derive(Default)]
pub(super) struct SourceI32Regions {
    // One byte per encoded byte; facts are indexed by instruction start.
    // This avoids a tree node and a heap allocation per numeric operator.
    // Bits: raw result, input present, binary, raw left, raw right.
    facts: Vec<u8>,
    pub value_strings: BTreeSet<usize>,
}

impl SourceI32Regions {
    pub fn is_raw(&self, at: &usize) -> bool {
        self.facts.get(*at).is_some_and(|facts| facts & 1 != 0)
    }

    pub fn inputs(&self, at: &usize) -> Option<&'static [bool]> {
        match self.facts.get(*at).copied().unwrap_or(0) & !1 {
            2 => Some(&[false]),
            10 => Some(&[true]),
            6 => Some(&[false, false]),
            14 => Some(&[true, false]),
            22 => Some(&[false, true]),
            30 => Some(&[true, true]),
            _ => None,
        }
    }

    pub fn new(
        chunk: &Chunk,
        root: &Chunk,
        ctx: &WasmTypeContext,
        imports: &std::collections::HashMap<(String, String, Option<u8>), usize>,
    ) -> Self {
        let mut plan = Self::default();
        let source = super::super::callable::source(chunk);
        let params: Vec<_> = chunk.func_sig.as_ref().map_or_else(Vec::new, |(p, _)| {
            p.split(',')
                .filter(|s| !s.is_empty())
                .map(super::super::types::val_type_byte)
                .collect()
        });
        let mut stack: Vec<Option<usize>> = Vec::new();
        let mut at = 0;
        let mut previous = None;
        while at < chunk.code.len() {
            let op = Op::from_code_at(&chunk.code, at);
            if op == Op::STRING_MEASURE_WTF8
                && let Some((Op::LOCAL_GET, prev)) = previous
            {
                let slot = u16::from_be_bytes([chunk.code[prev + 4], chunk.code[prev + 5]]);
                let ty = if (slot as usize) < params.len() {
                    Some(params[slot as usize])
                } else {
                    declared_scratch_type(chunk, slot as u32)
                };
                if !chunk.i64_locals.contains(&slot) && ty.is_none_or(|t| t == TYPE_EXTERNREF) {
                    plan.value_strings.insert(at);
                }
            }
            previous = Some((op, at));
            if !source {
                at += opcode_size(op, &chunk.code, at);
                continue;
            }
            if op == Op::LOCAL_GET {
                stack.push(None);
            } else if op == Op::LOCAL_SET || op == Op::LOCAL_TEE {
                // Only the stored top value crosses the boxed local ABI.
                // Numeric results below it are already evaluated and may
                // remain raw until their own operator consumes them.
                stack.pop();
                if op == Op::LOCAL_TEE {
                    stack.push(None);
                }
            } else if op == Op::CALL && !chunk.direct_defined_calls.contains_key(&at) {
                let index = u16::from_be_bytes([chunk.code[at + 4], chunk.code[at + 5]]) as usize;
                let argc = chunk.code[at + 6];
                let target = chunk
                    .imports
                    .get(index)
                    .or_else(|| root.imports.get(index))
                    .and_then(|import| {
                        imports
                            .get(&(import.module.clone(), import.name.clone(), Some(argc)))
                            .or_else(|| {
                                imports.get(&(import.module.clone(), import.name.clone(), None))
                            })
                    });
                let exact_one = target.is_some_and(|&target| {
                    ctx.func_result_arity.get(&(target as u32)) == Some(&1)
                        && ctx
                            .import_param_funcs
                            .get(&(target as u32))
                            .is_some_and(|p| p.len() == argc as usize)
                });
                if exact_one {
                    // Arguments retain their original boxing and evaluation
                    // order. The call cannot inspect numeric values below
                    // its arguments, and its result starts in the value ABI.
                    for _ in 0..argc {
                        stack.pop();
                    }
                    stack.push(None);
                } else {
                    stack.clear();
                }
            } else if op == Op::I32_CONST {
                stack.push(Some(at));
            } else if op.group() == 0 && matches!(op.sub(), 0x45..=0x4f | 0x67..=0x78) {
                let arity = if matches!(op.sub(), 0x45 | 0x67..=0x69) {
                    1
                } else {
                    2
                };
                if plan.facts.is_empty() {
                    plan.facts.resize(chunk.code.len(), 0);
                }
                let mut facts = 2 | if arity == 2 { 4 } else { 0 };
                for operand in (0..arity).rev() {
                    if let Some(producer) = stack.pop().flatten() {
                        plan.facts[producer] |= 1;
                        facts |= 8 << operand;
                    }
                }
                plan.facts[at] |= facts;
                stack.push(Some(at));
            } else {
                // No facts cross unknown calls, control-flow edges, or a
                // different value lane. Residual results retain the value ABI.
                stack.clear();
            }
            at += opcode_size(op, &chunk.code, at);
        }
        plan
    }
}
