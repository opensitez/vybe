//! Owned, bounded preparation for compiler-to-VM handoff.
//!
//! This module does not attach caches to mutable public Chunk fields. Its
//! immutable owner must be consumed by the VM or converted back to raw chunks.
use crate::vm::{BlockTargets, BrTableTargets};
use crate::{Chunk, Op, VM, VMError};
use std::mem::size_of;

pub(crate) struct CodeMetadata {
    pub(crate) ops: Vec<Op>,
    pub(crate) operands: Vec<u32>,
    pub(crate) wide: Vec<u64>,
    pub(crate) has_branch_table: bool,
    pub(crate) br_tables: Vec<Option<BrTableTargets>>,
    pub(crate) blocks: Vec<(usize, BlockTargets)>,
    pub(crate) calls: Vec<(usize, usize)>,
    pub(crate) globals: Vec<(usize, u32)>,
    pub(crate) raw_globals: Option<Vec<(usize, u32)>>,
    pub(crate) refs: Vec<(usize, u16)>,
}

/// Final emitted chunks plus context-free link metadata, with no mutable view.
///
/// Preparation is opt in. Budget rejection returns the original raw chunks;
/// VM validation failures are retained until the VM's normal validation step.
/// The caller owns the aggregate admission budget across pending workers.
pub struct PreparedLinkedCode {
    chunks: Vec<Chunk>,
    metadata: Result<Vec<CodeMetadata>, VMError>,
    metadata_bytes: usize,
}

impl PreparedLinkedCode {
    /// Prepare final owned code under a conservative retained-memory bound.
    /// `Err` means raw fallback, not a compilation or validation failure.
    pub fn prepare(chunks: Vec<Chunk>, max_metadata_bytes: usize) -> Result<Self, Vec<Chunk>> {
        // Dense buffers have an exact code-derived size. Sparse records and
        // optional structured caches reserve the remaining budget on demand.
        let mut bound = chunks.len().checked_mul(size_of::<CodeMetadata>());
        for chunk in &chunks {
            bound = bound.and_then(|n| {
                chunk
                    .code
                    .len()
                    .div_ceil(4)
                    .checked_mul(size_of::<Op>())
                    .and_then(|ops| {
                        chunk
                            .code
                            .len()
                            .div_ceil(4)
                            .checked_mul(size_of::<u32>())
                            .and_then(|operands| ops.checked_add(operands))
                    })
                    .and_then(|add| n.checked_add(add))
            });
        }
        let Some(used) = bound.filter(|n| *n <= max_metadata_bytes) else {
            return Err(chunks);
        };
        let mut budget = MetadataBudget {
            used,
            max: max_metadata_bytes,
        };
        // Validate direct-call headers and instruction shapes before preparation
        // can encounter raw-decoder panics. Unsafe-to-prepare shapes stay raw;
        // neither their error phase nor their existing decoder behavior changes.
        let scanned = (|| {
            let mut metadata = Vec::with_capacity(chunks.len());
            for chunk in &chunks {
                VM::validate_direct_defined_calls(chunk, &chunks)?;
                metadata.push(scan(chunk, &mut budget)?);
            }
            Ok(metadata)
        })();
        let metadata = match scanned {
            Ok(rows) => Ok(rows),
            Err(PreparationFailure::Raw) => return Err(chunks),
            Err(PreparationFailure::Invalid(error)) => Err(error),
        };
        let metadata_bytes = metadata.as_ref().map_or_else(
            |error| error.message.capacity(),
            |rows| {
                rows.capacity() * size_of::<CodeMetadata>()
                    + rows
                        .iter()
                        .map(|m| {
                            m.ops.capacity() * size_of::<Op>()
                                + m.operands.capacity() * size_of::<u32>()
                                + m.wide.capacity() * size_of::<u64>()
                                + m.br_tables.capacity() * size_of::<Option<BrTableTargets>>()
                                + m.blocks.capacity() * size_of::<(usize, BlockTargets)>()
                                + m.calls.capacity() * size_of::<(usize, usize)>()
                                + m.globals.capacity() * size_of::<(usize, u32)>()
                                + m.raw_globals
                                    .as_ref()
                                    .map_or(0, |g| g.capacity() * size_of::<(usize, u32)>())
                                + m.refs.capacity() * size_of::<(usize, u16)>()
                                + m.br_tables
                                    .iter()
                                    .flatten()
                                    .map(|t| t.depths.capacity() * size_of::<usize>())
                                    .sum::<usize>()
                        })
                        .sum::<usize>()
            },
        );
        if metadata_bytes > max_metadata_bytes {
            return Err(chunks);
        }
        Ok(Self {
            chunks,
            metadata,
            metadata_bytes,
        })
    }

    /// Shared bindings and standard-Wasm writers may inspect final code.
    pub fn chunks(&self) -> &[Chunk] {
        &self.chunks
    }

    /// Retained metadata bytes; excludes the owned code and allocator headers.
    pub fn metadata_bytes(&self) -> usize {
        self.metadata_bytes
    }

    /// Explicitly discard all trusted preparation before mutable raw use.
    pub fn into_raw(self) -> Vec<Chunk> {
        self.chunks
    }

    pub(crate) fn into_parts(self) -> (Vec<Chunk>, Result<Vec<CodeMetadata>, VMError>) {
        (self.chunks, self.metadata)
    }
}

enum PreparationFailure {
    Raw,
    Invalid(VMError),
}
impl From<VMError> for PreparationFailure {
    fn from(error: VMError) -> Self {
        Self::Invalid(error)
    }
}

struct MetadataBudget {
    used: usize,
    max: usize,
}
impl MetadataBudget {
    fn charge(&mut self, bytes: usize) -> Result<(), PreparationFailure> {
        let Some(next) = self.used.checked_add(bytes).filter(|n| *n <= self.max) else {
            return Err(PreparationFailure::Raw);
        };
        self.used = next;
        Ok(())
    }

    fn push<T>(&mut self, values: &mut Vec<T>, value: T) -> Result<(), PreparationFailure> {
        if values.len() == values.capacity() {
            let old = values.capacity();
            let Some(next) = old.checked_mul(2).map(|n| n.max(4)) else {
                return Err(PreparationFailure::Raw);
            };
            self.charge(
                (next - old)
                    .checked_mul(size_of::<T>())
                    .ok_or(PreparationFailure::Raw)?,
            )?;
            values.reserve_exact(next - old);
        }
        values.push(value);
        Ok(())
    }
}

fn scan(chunk: &Chunk, budget: &mut MetadataBudget) -> Result<CodeMetadata, PreparationFailure> {
    let code = &chunk.code;
    let mut m = CodeMetadata {
        ops: vec![Op::NOP; code.len().div_ceil(4)],
        operands: vec![0; code.len().div_ceil(4)],
        wide: Vec::new(),
        has_branch_table: false,
        br_tables: Vec::new(),
        blocks: Vec::new(),
        calls: Vec::new(),
        globals: Vec::new(),
        raw_globals: None,
        refs: Vec::new(),
    };
    let mut ip = 0usize;
    let mut void_start = 0usize;
    let mut void_len = 0usize;
    let mut direct_starts = 0usize;
    let mut direct_offsets = chunk.direct_defined_calls.keys().copied().peekable();
    let mut raw_walk_matches = true;
    let mut block_entries_bound = 0usize;
    let mut br_depths = 0usize;
    while ip + 3 < code.len() {
        let start = ip;
        // Ignore skipped payload offsets until the original end-of-scan
        // boundary check, preserving validation error precedence.
        while direct_offsets.peek().is_some_and(|&at| at < start) {
            direct_offsets.next();
        }
        let is_direct = direct_offsets.peek().copied() == Some(start);
        if is_direct {
            direct_starts += 1;
            direct_offsets.next();
        }
        let group = u16::from_be_bytes([code[ip], code[ip + 1]]);
        let sub = u16::from_be_bytes([code[ip + 2], code[ip + 3]]);
        let op=Op::decode(group,sub).ok_or_else(||VMError::new(format!(
            "invalid opcode 0x{:04X} 0x{:04X} at offset {} in chunk '{}' — module rejected at load",
            group,sub,ip,chunk.name)))?;
        let operand_start = start + 4;
        m.ops[start / 4] = op;
        m.operands[start / 4] = VM::mark_direct_call_operand(op, VM::decode_operand32(code, op, operand_start), is_direct);
        if op == Op::BLOCK && operand_start + 2 <= code.len() && m.operands[start / 4] == 0 {
            if void_len == 0 { void_start = start; }
            void_len += 1;
        } else if void_len != 0 {
            VM::finish_void_block_run(&mut m.operands, void_start, void_len);
            void_len = 0;
        }
        m.has_branch_table |= op == Op::BR_TABLE;
        // Wide values are charged on demand, not as a dense per-bucket row.
        // The paired u32 operand stores the pool index for these opcodes only.
        if op == Op::F64_CONST || op == Op::REF_FUNC {
            m.operands[start / 4] = u32::try_from(m.wide.len())
                .map_err(|_| PreparationFailure::Raw)?;
            budget.push(&mut m.wide, 0)?;
        }
        if op == Op::F64_CONST && operand_start + 7 < code.len() {
            *m.wide.last_mut().unwrap() =
                u64::from_le_bytes(code[operand_start..operand_start + 8].try_into().unwrap());
        }
        if op == Op::CALL && !is_direct {
            budget.push(
                &mut m.calls,
                (start, (m.operands[start / 4] & 0xffff) as usize),
            )?;
        }
        if (op == Op::GLOBAL_GET || op == Op::GLOBAL_SET) && operand_start + 3 < code.len() {
            budget.push(
                &mut m.globals,
                (
                    start,
                    u32::from_be_bytes(code[operand_start..operand_start + 4].try_into().unwrap()),
                ),
            )?;
        }
        if op == Op::REF_FUNC {
            if operand_start + 1 >= code.len() {
                return Err(PreparationFailure::Raw);
            }
            let relative = u16::from_be_bytes([code[operand_start], code[operand_start + 1]]);
            budget.push(&mut m.refs, (start, relative))?;
            m.wide[m.operands[start / 4] as usize] = relative as u64;
            ip += 4 + 2 + 1;
            if ip - 1 < code.len() {
                ip += (code[ip - 1] & 0x7f) as usize * 3;
            }
            raw_walk_matches &=
                ip == operand_start + op.operand_format().size_in(code, operand_start);
            continue;
        }
        if matches!(op, Op::BLOCK | Op::LOOP | Op::IF | Op::TRY_TABLE | Op::ELSE) {
            block_entries_bound += 1;
        }
        if op == Op::BR_TABLE {
            let mut count_ip = operand_start;
            let count = crate::opcode::read_leb_u32(code, &mut count_ip) as usize;
            // A malformed enormous count must not allocate unbounded metadata
            // on a compiler worker. Let the original raw loader handle it.
            if count > code.len() {
                return Err(PreparationFailure::Raw);
            }
            br_depths = br_depths
                .checked_add(count)
                .ok_or(PreparationFailure::Raw)?;
        }
        ip = operand_start + op.operand_format().size_in(code, operand_start);
    }
    VM::finish_void_block_run(&mut m.operands, void_start, void_len);
    if direct_starts != chunk.direct_defined_calls.len() {
        return Err(
            VMError::new("direct defined metadata is not on an instruction boundary").into(),
        );
    }
    if block_entries_bound != 0 {
        let capacity = block_entries_bound
            .checked_mul(2)
            .ok_or(PreparationFailure::Raw)?
            .max(4);
        budget.charge(
            capacity
                .checked_mul(size_of::<(usize, BlockTargets)>())
                .ok_or(PreparationFailure::Raw)?,
        )?;
    }
    m.blocks = crate::dispatch::build_block_entries(code);
    if m.has_branch_table {
        budget.charge(
            code.len()
                .checked_mul(size_of::<Option<BrTableTargets>>())
                .ok_or(PreparationFailure::Raw)?,
        )?;
        budget.charge(
            br_depths
                .checked_mul(size_of::<usize>())
                .ok_or(PreparationFailure::Raw)?,
        )?;
        m.br_tables = VM::build_br_table_targets_for_chunk(chunk);
    }
    if !raw_walk_matches {
        m.raw_globals = Some(raw_global_offsets(chunk, budget)?);
    }
    Ok(m)
}

// Exact original merge_global_table advancement, including capture-payload
// starts. This is a worker certificate of behavior, not a corrected decoder.
fn raw_global_offsets(
    chunk: &Chunk,
    budget: &mut MetadataBudget,
) -> Result<Vec<(usize, u32)>, PreparationFailure> {
    let code = &chunk.code;
    let mut records = Vec::new();
    let mut ip = 0usize;
    while ip + 3 < code.len() {
        let group = ((code[ip] as u16) << 8) | code[ip + 1] as u16;
        let sub = ((code[ip + 2] as u16) << 8) | code[ip + 3] as u16;
        let Some(op) = crate::opcode::Op::decode(group, sub) else {
            ip += 4;
            continue;
        };
        let operand_start = ip + 4;
        // Avoid unbounded work from a phantom branch-table header in capture
        // bytes. It is a raw-fallback admission decision, not a new VM error.
        if op == Op::BR_TABLE {
            let mut at = operand_start;
            if crate::opcode::read_leb_u32(code, &mut at) as usize > code.len() {
                return Err(PreparationFailure::Raw);
            }
        }
        let operand_len = op.operand_format().size_in(code, operand_start);
        if (op == crate::opcode::Op::GLOBAL_GET || op == crate::opcode::Op::GLOBAL_SET)
            && operand_start + 3 < code.len()
        {
            let old =
                u32::from_be_bytes(code[operand_start..operand_start + 4].try_into().unwrap());
            budget.push(&mut records, (ip, old))?;
        }
        ip = operand_start + operand_len;
    }
    Ok(records)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn preparation_omits_unused_wide_rows_and_keeps_wide_payloads() {
        let mut narrow = Chunk::new("narrow");
        narrow.emit_op_u16(Op::LOCAL_GET, 17, 0);
        narrow.emit_op(Op::RETURN, 0);
        let mut floating = narrow.clone();
        let at = floating.code.len();
        floating.emit_f64_const(f64::from_bits(0xfff8_0000_0000_0123), 0);
        let prepared = PreparedLinkedCode::prepare(vec![narrow, floating], usize::MAX).ok().unwrap();
        let (_, metadata) = prepared.into_parts();
        let metadata = metadata.ok().unwrap();
        assert_eq!(metadata[0].wide.capacity(), 0);
        assert_eq!(metadata[1].wide.len(), 1);
        assert_eq!(metadata[1].wide[metadata[1].operands[at / 4] as usize], 0xfff8_0000_0000_0123);
    }

    #[test]
    fn budget_rejection_and_explicit_raw_conversion_preserve_owned_code() {
        let mut chunk = Chunk::new("owned");
        chunk.emit_f64_const(7.0, 19);
        chunk.emit_op(Op::RETURN, 19);
        let expected = format!("{chunk:?}");
        let raw = PreparedLinkedCode::prepare(vec![chunk.clone()], 0)
            .err()
            .unwrap();
        assert_eq!(format!("{:?}", raw[0]), expected);
        let sealed = PreparedLinkedCode::prepare(vec![chunk], usize::MAX)
            .ok()
            .unwrap();
        assert!(sealed.metadata_bytes() > 0);
        assert_eq!(format!("{:?}", sealed.into_raw()[0]), expected);
    }

    #[test]
    fn absent_branch_tables_do_not_reserve_dense_branch_metadata() {
        let mut chunk = Chunk::new("no-branches");
        for _ in 0..4096 {
            chunk.code.extend_from_slice(&Op::NOP.encode());
        }
        let seal = PreparedLinkedCode::prepare(vec![chunk], 1024 * 1024)
            .ok()
            .unwrap();
        assert!(seal.metadata_bytes() < 256 * 1024);
    }

    #[test]
    fn validation_failure_remains_staged_with_original_chunks() {
        let mut chunk = Chunk::new("bad");
        chunk.code.extend_from_slice(&[255, 255, 255, 255]);
        let sealed = PreparedLinkedCode::prepare(vec![chunk], usize::MAX)
            .ok()
            .unwrap();
        assert_eq!(sealed.chunks()[0].code, [255, 255, 255, 255]);
        let (_, metadata) = sealed.into_parts();
        assert!(
            metadata
                .err()
                .unwrap()
                .message
                .contains("module rejected at load")
        );
    }

    #[test]
    fn truncated_reference_stays_on_raw_loader_path() {
        let mut chunk = Chunk::new("short-ref");
        chunk.emit_op(Op::REF_FUNC, 0);
        assert!(PreparedLinkedCode::prepare(vec![chunk], usize::MAX).is_err());
    }
    #[test]
    fn direct_site_cursor_preserves_host_call_classification() {
        let mut caller = Chunk::new("caller");
        for _ in 0..3 {
            caller.emit_call(0, 0, 0);
        }
        caller.direct_defined_calls.insert(7, crate::chunk::DirectDefinedCall {
            target: 1, source: true, params: String::new(), results: "externref".into(),
        });
        let mut helper = Chunk::new("helper");
        helper.emit_op(Op::NULL, 0);
        helper.emit_op(Op::RETURN, 0);
        let seal = PreparedLinkedCode::prepare(vec![caller, helper], usize::MAX).ok().unwrap();
        let (_, metadata) = seal.into_parts();
        assert_eq!(metadata.ok().unwrap()[0].calls, vec![(0, 0), (14, 0)]);
    }

    #[test]
    fn skipped_direct_payload_offset_retains_error_precedence() {
        let mut caller = Chunk::new("caller");
        caller.emit_f64_const(0.0, 0);
        caller.code[4..8].copy_from_slice(&Op::CALL.encode());
        caller.code[8..11].fill(0);
        caller.direct_defined_calls.insert(4, crate::chunk::DirectDefinedCall {
            target: 1, source: true, params: String::new(), results: "externref".into(),
        });
        let mut helper = Chunk::new("helper");
        helper.emit_op(Op::NULL, 0);
        helper.emit_op(Op::RETURN, 0);
        for invalid_opcode in [false, true] {
            let mut chunk = caller.clone();
            if invalid_opcode {
                chunk.code.extend_from_slice(&[0xab, 0xcd, 0, 0]);
            }
            let seal = PreparedLinkedCode::prepare(vec![chunk, helper.clone()], usize::MAX).ok().unwrap();
            let (_, metadata) = seal.into_parts();
            let message = metadata.err().unwrap().message;
            if invalid_opcode {
                assert!(message.contains("invalid opcode"));
            } else {
                assert_eq!(message, "direct defined metadata is not on an instruction boundary");
            }
        }
    }

}
