//! Typed scalar core validation. Coverage is explicit until proposal validators
//! have been migrated; an unhandled declaration/instruction is never accepted.
use super::{WasmError, section_types};
use crate::binary_cursor::Cursor;
use crate::binary_sections::{BinaryKind, Sections};
use vybe_runtime::chunk::NativeValueType;

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(super) enum Coverage {
    Complete,
    RequiresProposalValidation,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum Type {
    I32,
    I64,
    F32,
    F64,
    V128,
}
use Type::*;

fn scalar(cursor: &mut Cursor<'_>) -> Result<Option<Type>, WasmError> {
    Ok(match section_types::value_type(cursor)? {
        NativeValueType::I32 => Some(I32),
        NativeValueType::I64 => Some(I64),
        NativeValueType::F32 => Some(F32),
        NativeValueType::F64 => Some(F64),
        NativeValueType::V128 => Some(V128),
        NativeValueType::Ref(_) => None,
    })
}

#[derive(Default)]
struct Signature {
    params: Vec<Type>,
    results: Vec<Type>,
}

#[derive(Clone, Copy, PartialEq, Eq)]
enum Kind {
    Function,
    Block,
    Loop,
    If,
    Else,
}

struct Frame<'a> {
    kind: Kind,
    height: usize,
    params: &'a [Type],
    results: &'a [Type],
    unreachable: bool,
}

struct Stack<'a> {
    values: Vec<Option<Type>>,
    frames: Vec<Frame<'a>>,
}

fn invalid(message: &'static str) -> WasmError {
    WasmError::invalid(message)
}

impl<'a> Stack<'a> {
    fn new(results: &'a [Type]) -> Self {
        Self {
            values: Vec::new(),
            frames: vec![Frame {
                kind: Kind::Function,
                height: 0,
                params: &[],
                results,
                unreachable: false,
            }],
        }
    }
    fn pop(&mut self) -> Result<Option<Type>, WasmError> {
        let frame = self
            .frames
            .last()
            .ok_or_else(|| invalid("instruction after function end"))?;
        if self.values.len() == frame.height {
            return if frame.unreachable {
                Ok(None)
            } else {
                Err(invalid("operand stack underflow"))
            };
        }
        self.values
            .pop()
            .ok_or_else(|| invalid("operand stack underflow"))
    }
    fn pop_type(&mut self, expected: Type) -> Result<(), WasmError> {
        if self.pop()?.is_some_and(|actual| actual != expected) {
            return Err(invalid("operand type mismatch"));
        }
        Ok(())
    }
    fn pop_types(&mut self, types: &[Type]) -> Result<(), WasmError> {
        for &ty in types.iter().rev() {
            self.pop_type(ty)?;
        }
        Ok(())
    }
    fn push_types(&mut self, types: &[Type]) {
        self.values.extend(types.iter().copied().map(Some));
    }
    fn effect(&mut self, inputs: &[Type], outputs: &[Type]) -> Result<(), WasmError> {
        self.pop_types(inputs)?;
        self.push_types(outputs);
        Ok(())
    }
    fn unreachable(&mut self) {
        let frame = self.frames.last_mut().unwrap();
        self.values.truncate(frame.height);
        frame.unreachable = true;
    }
    fn enter(
        &mut self,
        kind: Kind,
        params: &'a [Type],
        results: &'a [Type],
    ) -> Result<(), WasmError> {
        self.pop_types(params)?;
        self.frames.push(Frame {
            kind,
            height: self.values.len(),
            params,
            results,
            unreachable: false,
        });
        self.push_types(params);
        Ok(())
    }
    fn finish_arm(&mut self) -> Result<(), WasmError> {
        let frame = self.frames.last().unwrap();
        let (height, results) = (frame.height, frame.results);
        self.pop_types(results)?;
        if self.values.len() != height {
            return Err(invalid("block leaves extra operands"));
        }
        Ok(())
    }
    fn else_(&mut self) -> Result<(), WasmError> {
        if self.frames.last().unwrap().kind != Kind::If {
            return Err(invalid("unexpected else"));
        }
        self.finish_arm()?;
        let frame = self.frames.last_mut().unwrap();
        frame.kind = Kind::Else;
        frame.unreachable = false;
        let params = frame.params;
        self.push_types(params);
        Ok(())
    }
    fn end(&mut self) -> Result<(), WasmError> {
        let frame = self.frames.last().unwrap();
        if frame.kind == Kind::If && frame.params != frame.results {
            return Err(invalid("if without else has incompatible implicit arm"));
        }
        self.finish_arm()?;
        let frame = self.frames.pop().unwrap();
        self.push_types(frame.results);
        Ok(())
    }
    fn label(&self, depth: u32) -> Result<&'a [Type], WasmError> {
        let index = self
            .frames
            .len()
            .checked_sub(depth as usize + 1)
            .ok_or_else(|| invalid("unknown branch label"))?;
        let frame = &self.frames[index];
        Ok(if frame.kind == Kind::Loop {
            frame.params
        } else {
            frame.results
        })
    }
}

fn numeric(op: u8) -> Option<(&'static [Type], &'static [Type])> {
    Some(match op {
        0x45 => (&[I32], &[I32]),
        0x46..=0x4f => (&[I32, I32], &[I32]),
        0x50 => (&[I64], &[I32]),
        0x51..=0x5a => (&[I64, I64], &[I32]),
        0x5b..=0x60 => (&[F32, F32], &[I32]),
        0x61..=0x66 => (&[F64, F64], &[I32]),
        0x67..=0x69 | 0xc0..=0xc1 => (&[I32], &[I32]),
        0x6a..=0x78 => (&[I32, I32], &[I32]),
        0x79..=0x7b | 0xc2..=0xc4 => (&[I64], &[I64]),
        0x7c..=0x8a => (&[I64, I64], &[I64]),
        0x8b..=0x91 => (&[F32], &[F32]),
        0x92..=0x98 => (&[F32, F32], &[F32]),
        0x99..=0x9f => (&[F64], &[F64]),
        0xa0..=0xa6 => (&[F64, F64], &[F64]),
        0xa7 => (&[I64], &[I32]),
        0xa8..=0xa9 => (&[F32], &[I32]),
        0xaa..=0xab => (&[F64], &[I32]),
        0xac..=0xad => (&[I32], &[I64]),
        0xae..=0xaf => (&[F32], &[I64]),
        0xb0..=0xb1 => (&[F64], &[I64]),
        0xb2..=0xb3 => (&[I32], &[F32]),
        0xb4..=0xb5 => (&[I64], &[F32]),
        0xb6 => (&[F64], &[F32]),
        0xb7..=0xb8 => (&[I32], &[F64]),
        0xb9..=0xba => (&[I64], &[F64]),
        0xbb => (&[F32], &[F64]),
        0xbc => (&[F32], &[I32]),
        0xbd => (&[F64], &[I64]),
        0xbe => (&[I32], &[F32]),
        0xbf => (&[I64], &[F64]),
        _ => return None,
    })
}

fn block_signature<'a>(
    cursor: &mut Cursor<'_>,
    types: &'a [Signature],
) -> Result<Option<(&'a [Type], &'a [Type])>, WasmError> {
    if cursor.peek() == Some(0x40) {
        cursor.byte()?;
        return Ok(Some((&[], &[])));
    }
    if matches!(cursor.peek(), Some(0x63..=0x65 | 0x68..=0x75)) {
        return Ok(None);
    }
    if matches!(cursor.peek(), Some(0x7b..=0x7f)) {
        let results: &[Type] = match scalar(cursor)?.unwrap() {
            I32 => &[I32],
            I64 => &[I64],
            F32 => &[F32],
            F64 => &[F64],
            V128 => &[V128],
        };
        return Ok(Some((&[], results)));
    }
    let index = u32::try_from(cursor.s33()?).map_err(|_| WasmError::from("invalid block type"))?;
    let sig = types
        .get(index as usize)
        .ok_or_else(|| invalid("unknown block type"))?;
    Ok(Some((&sig.params, &sig.results)))
}

// Run-length declarations avoid expanding a large local group into a huge Vec.
struct Locals<'a> {
    params: &'a [Type],
    groups: Vec<(u32, Type)>,
}
impl Locals<'_> {
    fn get(&self, index: u32) -> Result<Type, WasmError> {
        if let Some(ty) = self.params.get(index as usize) {
            return Ok(*ty);
        }
        let index = index - self.params.len() as u32;
        self.groups
            .get(self.groups.partition_point(|&(end, _)| end <= index))
            .map(|&(_, ty)| ty)
            .ok_or_else(|| invalid("unknown local"))
    }
}

fn function(
    body: &[u8],
    sig: &Signature,
    types: &[Signature],
    functions: &[u32],
) -> Result<Coverage, WasmError> {
    let mut cursor = Cursor::new(body);
    let count = section_types::count(&mut cursor, 2)?;
    let mut locals = Locals {
        params: &sig.params,
        groups: Vec::with_capacity(count),
    };
    let mut total = 0u32;
    for _ in 0..count {
        let count = cursor.u32()?;
        let Some(ty) = scalar(&mut cursor)? else {
            return Ok(Coverage::RequiresProposalValidation);
        };
        total = total
            .checked_add(count)
            .filter(|&n| u64::from(n) + sig.params.len() as u64 <= u64::from(u32::MAX))
            .ok_or_else(|| invalid("too many locals"))?;
        locals.groups.push((total, ty));
    }
    let mut stack = Stack::new(&sig.results);
    while cursor.remaining() != 0 {
        if stack.frames.is_empty() {
            return Err("instructions after function end".into());
        }
        let op = cursor.byte()?;
        if let Some((inputs, outputs)) = numeric(op) {
            stack.effect(inputs, outputs)?;
            continue;
        }
        match op {
            0x00 => stack.unreachable(),
            0x01 => {}
            0x02..=0x04 => {
                if op == 0x04 {
                    stack.pop_type(I32)?;
                }
                let Some((params, results)) = block_signature(&mut cursor, types)? else {
                    return Ok(Coverage::RequiresProposalValidation);
                };
                stack.enter(
                    match op {
                        2 => Kind::Block,
                        3 => Kind::Loop,
                        _ => Kind::If,
                    },
                    params,
                    results,
                )?;
            }
            0x05 => stack.else_()?,
            0x0b => stack.end()?,
            0x0c | 0x0d => {
                let target = stack.label(cursor.u32()?)?;
                if op == 0x0d {
                    stack.pop_type(I32)?;
                }
                stack.pop_types(target)?;
                if op == 0x0d {
                    stack.push_types(target);
                } else {
                    stack.unreachable();
                }
            }
            0x0e => {
                let count = section_types::count(&mut cursor, 1)?;
                let mut target: Option<&[Type]> = None;
                for _ in 0..=count {
                    let label = stack.label(cursor.u32()?)?;
                    if target.is_some_and(|target| target != label) {
                        return Err(invalid("branch table label types differ"));
                    }
                    target = Some(label);
                }
                stack.pop_type(I32)?;
                stack.pop_types(target.unwrap())?;
                stack.unreachable();
            }
            0x0f => {
                stack.pop_types(&sig.results)?;
                stack.unreachable();
            }
            0x10 | 0x12 => {
                let index = cursor.u32()?;
                let type_index = *functions
                    .get(index as usize)
                    .ok_or_else(|| invalid("unknown function"))?;
                let callee = &types[type_index as usize];
                stack.pop_types(&callee.params)?;
                if op == 0x12 {
                    if callee.results != sig.results {
                        return Err(invalid("tail call result types differ"));
                    }
                    stack.unreachable();
                } else {
                    stack.push_types(&callee.results);
                }
            }
            0x1a => {
                stack.pop()?;
            }
            0x1b => {
                stack.pop_type(I32)?;
                let rhs = stack.pop()?;
                let lhs = stack.pop()?;
                if lhs.is_some() && rhs.is_some() && lhs != rhs {
                    return Err(invalid("select operand types differ"));
                }
                stack.values.push(lhs.or(rhs));
            }
            0x1c => {
                if cursor.u32()? != 1 {
                    return Err(invalid("typed select requires one result"));
                }
                let Some(ty) = scalar(&mut cursor)? else {
                    return Ok(Coverage::RequiresProposalValidation);
                };
                stack.pop_type(I32)?;
                stack.effect(&[ty, ty], &[ty])?;
            }
            0x20..=0x22 => {
                let ty = locals.get(cursor.u32()?)?;
                if op != 0x20 {
                    stack.pop_type(ty)?;
                }
                if op != 0x21 {
                    stack.push_types(&[ty]);
                }
            }
            0x41 => {
                cursor.i32()?;
                stack.push_types(&[I32]);
            }
            0x42 => {
                cursor.i64()?;
                stack.push_types(&[I64]);
            }
            0x43 => {
                cursor.bytes(4)?;
                stack.push_types(&[F32]);
            }
            0x44 => {
                cursor.bytes(8)?;
                stack.push_types(&[F64]);
            }
            0xfc => {
                let (inputs, outputs): (&[Type], &[Type]) = match cursor.u32()? {
                    0 | 1 => (&[F32], &[I32]),
                    2 | 3 => (&[F64], &[I32]),
                    4 | 5 => (&[F32], &[I64]),
                    6 | 7 => (&[F64], &[I64]),
                    _ => return Ok(Coverage::RequiresProposalValidation),
                };
                stack.effect(inputs, outputs)?;
            }
            _ => return Ok(Coverage::RequiresProposalValidation),
        }
    }
    if !stack.frames.is_empty() {
        return Err("function body missing end".into());
    }
    Ok(Coverage::Complete)
}

pub(super) fn validate(data: &[u8]) -> Result<Coverage, WasmError> {
    let sections = Sections::new(data)?;
    if sections.kind() != BinaryKind::Module {
        return Ok(Coverage::RequiresProposalValidation);
    }
    let mut payloads = [None; 14];
    let mut last = 0;
    for section in sections {
        let section = section?;
        if !matches!(section.id, 0 | 1 | 2 | 3 | 7 | 8 | 10) {
            return Ok(Coverage::RequiresProposalValidation);
        }
        if section.id == 0 {
            Cursor::new(section.payload).name()?;
            continue;
        }
        let slot = &mut payloads[section.id as usize];
        if slot.is_some() || super::section_order_rank(section.id) < super::section_order_rank(last)
        {
            return Err("duplicate or out-of-order section".into());
        }
        *slot = Some(section.payload);
        last = section.id;
    }
    let mut types = Vec::new();
    if let Some(bytes) = payloads[1] {
        let mut cursor = Cursor::new(bytes);
        let count = section_types::count(&mut cursor, 1)?;
        for _ in 0..count {
            if cursor.byte()? != 0x60 {
                return Ok(Coverage::RequiresProposalValidation);
            }
            let mut sig = Signature::default();
            for target in [&mut sig.params, &mut sig.results] {
                let count = section_types::count(&mut cursor, 1)?;
                for _ in 0..count {
                    let Some(ty) = scalar(&mut cursor)? else {
                        return Ok(Coverage::RequiresProposalValidation);
                    };
                    target.push(ty);
                }
            }
            types.push(sig);
        }
        section_types::finish(&cursor)?;
    }
    let mut functions = Vec::new();
    if let Some(bytes) = payloads[2] {
        let mut covered = true;
        section_types::for_each_import(bytes, |_, _, import| {
            if let section_types::ImportType::Func {
                type_index,
                exact: false,
            } = import
            {
                if *type_index as usize >= types.len() {
                    return Err(invalid("unknown imported function type"));
                }
                functions.push(*type_index);
            } else {
                covered = false;
            }
            Ok(())
        })?;
        if !covered {
            return Ok(Coverage::RequiresProposalValidation);
        }
    }
    let imported_functions = functions.len();
    if let Some(bytes) = payloads[3] {
        let mut cursor = Cursor::new(bytes);
        let count = section_types::count(&mut cursor, 1)?;
        for _ in 0..count {
            let index = cursor.u32()?;
            if index as usize >= types.len() {
                return Err(invalid("unknown function type"));
            }
            functions.push(index);
        }
        section_types::finish(&cursor)?;
    }
    if let Some(bytes) = payloads[7] {
        let mut cursor = Cursor::new(bytes);
        let count = section_types::count(&mut cursor, 3)?;
        let mut names = std::collections::HashSet::with_capacity(count);
        for _ in 0..count {
            if !names.insert(cursor.name()?) {
                return Err(invalid("duplicate export name"));
            }
            if cursor.byte()? != 0 {
                return Ok(Coverage::RequiresProposalValidation);
            }
            if cursor.u32()? as usize >= functions.len() {
                return Err(invalid("unknown exported function"));
            }
        }
        section_types::finish(&cursor)?;
    }
    if let Some(bytes) = payloads[8] {
        let mut cursor = Cursor::new(bytes);
        let index = cursor.u32()?;
        section_types::finish(&cursor)?;
        let type_index = functions
            .get(index as usize)
            .ok_or_else(|| invalid("unknown start function"))?;
        let sig = &types[*type_index as usize];
        if !sig.params.is_empty() || !sig.results.is_empty() {
            return Err(invalid("start function must have empty signature"));
        }
    }
    let mut bodies = 0;
    if let Some(bytes) = payloads[10] {
        let mut cursor = Cursor::new(bytes);
        bodies = section_types::count(&mut cursor, 1)?;
        if bodies != functions.len() - imported_functions {
            return Err("function/code count mismatch".into());
        }
        for &index in &functions[imported_functions..] {
            let size = cursor.u32()? as usize;
            let body = cursor.bytes(size)?;
            if function(body, &types[index as usize], &types, &functions)? != Coverage::Complete {
                return Ok(Coverage::RequiresProposalValidation);
            }
        }
        section_types::finish(&cursor)?;
    }
    if bodies != functions.len() - imported_functions {
        return Err("function/code count mismatch".into());
    }
    Ok(Coverage::Complete)
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::encoding::*;

    fn byte(ty: Type) -> u8 {
        match ty {
            I32 => 0x7f,
            I64 => 0x7e,
            F32 => 0x7d,
            F64 => 0x7c,
            V128 => 0x7b,
        }
    }
    fn module(signatures: &[(&[u8], &[u8])], indices: &[u32], bodies: &[Vec<u8>]) -> Vec<u8> {
        let mut out = b"\0asm\x01\0\0\0".to_vec();
        let mut types = vec![];
        write_leb128_u32(&mut types, signatures.len() as u32);
        for &(params, results) in signatures {
            types.push(0x60);
            write_leb128_u32(&mut types, params.len() as u32);
            types.extend_from_slice(params);
            write_leb128_u32(&mut types, results.len() as u32);
            types.extend_from_slice(results);
        }
        write_section(&mut out, 1, &types);
        let mut functions = vec![];
        write_leb128_u32(&mut functions, indices.len() as u32);
        for &index in indices {
            write_leb128_u32(&mut functions, index);
        }
        write_section(&mut out, 3, &functions);
        let mut code = vec![];
        write_leb128_u32(&mut code, bodies.len() as u32);
        for body in bodies {
            write_leb128_u32(&mut code, body.len() as u32);
            code.extend_from_slice(body);
        }
        write_section(&mut out, 10, &code);
        out
    }
    fn oracle(bytes: &[u8]) -> bool {
        wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all())
            .validate_all(bytes)
            .is_ok()
    }
    fn compare(bytes: &[u8], expected: bool) {
        assert_eq!(oracle(bytes), expected, "reference fixture {bytes:?}");
        let actual = validate(bytes);
        if expected {
            assert_eq!(actual.unwrap(), Coverage::Complete, "{bytes:?}");
        } else {
            assert!(actual.is_err(), "accepted invalid fixture {bytes:?}");
        }
    }

    #[test]
    fn every_scalar_numeric_signature_checks_inputs_and_results() {
        for op in 0x45..=0xc4 {
            let (inputs, outputs) = numeric(op).unwrap();
            let params: Vec<_> = inputs.iter().copied().map(byte).collect();
            let results: Vec<_> = outputs.iter().copied().map(byte).collect();
            let mut body = vec![0];
            for i in 0..inputs.len() {
                body.extend_from_slice(&[0x20, i as u8]);
            }
            body.extend_from_slice(&[op, 0x0b]);
            compare(&module(&[(&params, &results)], &[0], &[body.clone()]), true);
            for i in 0..params.len() {
                for wrong in [I32, I64, F32, F64, V128] {
                    if wrong == inputs[i] {
                        continue;
                    }
                    let mut params = params.clone();
                    params[i] = byte(wrong);
                    compare(
                        &module(&[(&params, &results)], &[0], &[body.clone()]),
                        false,
                    );
                }
            }
            let wrong = [if outputs[0] == I32 { 0x7e } else { 0x7f }];
            compare(&module(&[(&params, &wrong)], &[0], &[body]), false);
        }
    }

    #[test]
    fn structured_control_uses_typed_labels_and_polymorphic_unreachable() {
        let signatures: &[(&[u8], &[u8])] = &[(&[], &[0x7f]), (&[0x7f], &[0x7f]), (&[], &[0x7e])];
        for (body, valid) in [
            (vec![0, 0x41, 7, 2, 1, 0x0b, 0x0b], true), // block param/result
            (vec![0, 0x42, 7, 2, 1, 0x0b, 0x0b], false),
            (vec![0, 0x41, 7, 0x41, 1, 4, 1, 0x0b, 0x0b], true), // implicit else
            (vec![0, 0x41, 1, 4, 0x7f, 0x41, 7, 0x0b, 0x0b], false),
            (
                vec![0, 0x41, 1, 4, 0x7f, 0x41, 7, 5, 0x41, 8, 0x0b, 0x0b],
                true,
            ),
            (
                vec![0, 0x41, 1, 4, 0x7f, 0x41, 7, 5, 0x42, 8, 0x0b, 0x0b],
                false,
            ),
            (vec![0, 0, 0x6a, 0x0b], true), // polymorphic operands
            (vec![0, 0, 0x42, 1, 0x6a, 0x0b], false), // concrete mismatched operand
            (vec![0, 0, 0x41, 1, 0x41, 2, 0x0b], false), // extra operand
            (
                vec![0, 2, 0x7f, 0x41, 7, 0x41, 1, 0x0d, 0, 0x0b, 0x0b],
                true,
            ),
            (
                vec![0, 2, 0x7f, 0x42, 7, 0x41, 1, 0x0d, 0, 0x0b, 0x0b],
                false,
            ),
            (vec![0, 2, 0x7f, 0x41, 7, 0x0c, 0, 0x0b, 0x0b], true),
            (vec![0, 0x0c, 1, 0x0b], false),
        ] {
            compare(&module(signatures, &[0], &[body]), valid);
        }
    }

    #[test]
    fn direct_and_tail_calls_check_declared_signatures() {
        for (body, valid) in [
            (vec![0, 0x41, 5, 0x10, 1, 0x0b], true),
            (vec![0, 0x42, 5, 0x10, 1, 0x0b], false),
            (vec![0, 0x41, 5, 0x12, 1, 0x0b], true),
            (vec![0, 0x10, 2, 0x0b], false),
        ] {
            compare(
                &module(
                    &[(&[], &[0x7f]), (&[0x7f], &[0x7f])],
                    &[0, 1],
                    &[body, vec![0, 0x20, 0, 0x0b]],
                ),
                valid,
            );
        }
        compare(
            &module(
                &[(&[], &[0x7f]), (&[], &[0x7e])],
                &[0, 1],
                &[vec![0, 0x12, 1, 0x0b], vec![0, 0x42, 1, 0x0b]],
            ),
            false,
        );
    }

    #[test]
    fn local_groups_are_compact_and_indices_are_checked() {
        for (body, valid) in [
            (vec![2, 2, 0x7e, 1, 0x7f, 0x20, 2, 0x0b], true),
            (vec![2, 2, 0x7e, 1, 0x7f, 0x20, 1, 0x0b], false),
            (vec![1, 0, 0x7f, 0x20, 0, 0x0b], false),
            (vec![1, 1, 0x7f, 0x41, 1, 0x22, 0, 0x0b], true),
            (vec![1, 1, 0x7f, 0x42, 1, 0x21, 0, 0x41, 1, 0x0b], false),
        ] {
            compare(&module(&[(&[], &[0x7f])], &[0], &[body]), valid);
        }
        let locals = Locals {
            params: &[I32],
            groups: vec![(1_000_000_000, I64), (1_000_000_001, F32)],
        };
        assert_eq!(locals.get(1_000_000_000).unwrap(), I64);
        assert_eq!(locals.get(1_000_000_001).unwrap(), F32);
        assert!(locals.get(1_000_000_002).is_err());
        assert_eq!(locals.groups.capacity(), 2);
    }

    #[test]
    fn proposal_coverage_is_explicit_not_silent_acceptance() {
        let reference_type = module(&[(&[0x6f], &[])], &[0], &[vec![0, 0x0b]]);
        assert_eq!(
            validate(&reference_type).unwrap(),
            Coverage::RequiresProposalValidation
        );
        let memory_instruction = module(&[(&[], &[0x7f])], &[0], &[vec![0, 0x3f, 0, 0x0b]]);
        assert_eq!(
            validate(&memory_instruction).unwrap(),
            Coverage::RequiresProposalValidation
        );
        assert!(!oracle(&memory_instruction));
    }

    #[test]
    fn branch_tables_and_select_keep_typed_control_contracts() {
        for (body, valid) in [
            (
                vec![0, 2, 0x7f, 0x41, 7, 0x41, 0, 0x0e, 1, 0, 1, 0x0b, 0x0b],
                true,
            ),
            (
                vec![
                    0, 2, 0x7e, 0x42, 7, 0x41, 0, 0x0e, 1, 0, 1, 0x0b, 0x1a, 0x41, 0, 0x0b,
                ],
                false,
            ),
            (vec![0, 0x41, 1, 0x41, 2, 0x41, 0, 0x1b, 0x0b], true),
            (vec![0, 0x41, 1, 0x42, 2, 0x41, 0, 0x1b, 0x0b], false),
            (
                vec![0, 0x41, 1, 0x41, 2, 0x41, 0, 0x1c, 1, 0x7f, 0x0b],
                true,
            ),
            (
                vec![0, 0x41, 1, 0x41, 2, 0x41, 0, 0x1c, 1, 0x7e, 0x0b],
                false,
            ),
            (vec![0, 0, 0x1b, 0x0b], true),
            (vec![0, 0x41, 1], false),
            (vec![0, 0x41, 1, 0x0b, 0x0b], false),
            (vec![0, 0x41, 1, 0x05, 0x0b], false),
        ] {
            compare(&module(&[(&[], &[0x7f])], &[0], &[body]), valid);
        }
    }

    #[test]
    fn generated_flat_scalar_programs_match_reference_validation() {
        let tokens: &[&[u8]] = &[
            &[0],
            &[1],
            &[0x41, 7],
            &[0x42, 7],
            &[0x43, 0, 0, 0, 0],
            &[0x44, 0, 0, 0, 0, 0, 0, 0, 0],
            &[0x1a],
            &[0x1b],
            &[0x6a],
            &[0x7c],
            &[0xa7],
            &[0xac],
            &[0x45],
            &[0x0f],
            &[0x0c, 0],
            &[0x0d, 0],
            &[0x20, 0],
            &[0x21, 0],
            &[0x22, 0],
        ];
        let mut seed = 0x6a09_e667u32;
        for _ in 0..2048 {
            seed = seed.wrapping_mul(1664525).wrapping_add(1013904223);
            let length = (seed % 12) as usize;
            let mut body = vec![0];
            for _ in 0..length {
                seed = seed.wrapping_mul(1664525).wrapping_add(1013904223);
                body.extend_from_slice(tokens[seed as usize % tokens.len()]);
            }
            body.push(0x0b);
            let bytes = module(&[(&[0x7f], &[0x7f])], &[0], &[body]);
            let reference = oracle(&bytes);
            let actual = validate(&bytes);
            assert_eq!(
                matches!(actual, Ok(Coverage::Complete)),
                reference,
                "{bytes:?}: {actual:?}"
            );
            assert_ne!(actual.ok(), Some(Coverage::RequiresProposalValidation));
        }
    }

    #[test]
    fn scalar_imports_preserve_index_spaces_and_compact_import_grammar() {
        fn with_imports(bytes: &[u8], imports: &[u8]) -> Vec<u8> {
            let mut out = bytes[..8].to_vec();
            for section in Sections::new(bytes).unwrap() {
                let section = section.unwrap();
                write_section(&mut out, section.id, section.payload);
                if section.id == 1 {
                    write_section(&mut out, 2, imports);
                }
            }
            out
        }
        let fixture = module(
            &[(&[0x7f], &[0x7f]), (&[], &[0x7f])],
            &[1],
            &[vec![0, 0x41, 5, 0x10, 0, 0x0b]],
        );
        for imports in [
            &[1, 1, b'm', 1, b'f', 0, 0][..],
            &[1, 1, b'm', 0, 0x7f, 1, 1, b'f', 0, 0],
            &[1, 1, b'm', 0, 0x7e, 0, 0, 1, 1, b'f'],
        ] {
            compare(&with_imports(&fixture, imports), true);
        }
        compare(&with_imports(&fixture, &[1, 1, b'm', 1, b'f', 0, 3]), false);
        let wrong = module(
            &[(&[0x7e], &[0x7f]), (&[], &[0x7f])],
            &[1],
            &[vec![0, 0x41, 5, 0x10, 0, 0x0b]],
        );
        compare(&with_imports(&wrong, &[1, 1, b'm', 1, b'f', 0, 0]), false);
    }
}
