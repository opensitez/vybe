//! First-entry preparation for the immutable PreparedLinkedCode carrier.
//! Validation and module-bound native target capture remain at installation.
//! This generates interpreter execution metadata, not source code or machine JIT.
use super::*;
impl VM {
    pub(super) fn prepare_first_entry_decode(&mut self, chunk_index: usize) {
        let Some(base) = self.pending_entry_decode.remove(&chunk_index) else { return; };
        let chunk = &self.chunks[chunk_index];
        let (ops, operands, wide, has_br_table) = Self::build_decoded_caches_with_branch_presence(chunk, base);
        let branches = if has_br_table { Self::build_br_table_targets_for_chunk(chunk) } else { Vec::new() };
        let blocks = Self::build_block_targets_for_chunk(chunk);
        self.decoded_ops[chunk_index] = ops;
        self.decoded_operands[chunk_index] = operands;
        self.decoded_operands64[chunk_index] = wide;
        self.br_table_targets[chunk_index] = branches;
        self.block_targets[chunk_index] = blocks;
    }
}

#[cfg(test)]
#[path="first_entry_preparation_tests.rs"]
mod tests;
