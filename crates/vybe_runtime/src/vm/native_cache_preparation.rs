//! Allocate dense native-site metadata only for an entered function body.

use super::VM;

impl VM {
    pub(super) fn prepare_native_call_sites_for_entry(&mut self, chunk_index: usize) {
        let Some(pending) = self.pending_native_call_sites.remove(&chunk_index) else {
            return;
        };
        let sites = &mut self.native_call_sites[chunk_index];
        if sites.is_empty() && !pending.is_empty() {
            sites.resize(self.chunks[chunk_index].code.len().div_ceil(4), None);
        }
        // Targets were resolved while their module's import table was active.
        // Do not consult a caller's restored table or transient lookup cache.
        for (offset, target) in pending {
            sites[offset / 4] = Some(target);
        }
    }
}

#[cfg(test)]
mod tests {
    use super::super::{NativeCallSiteTarget, Op};
    use super::*;
    use crate::chunk::Chunk;
    use crate::vm::{JsStringBuiltin, ResolvedCallTarget};

    fn fixture(fallback: bool) -> VM {
        let mut vm = VM::new();
        let mut chunk = Chunk::new("first-entry-native-cache");
        chunk.local_count = 1;
        let import = chunk.add_import("wasm:js-string", "length");
        assert_eq!(import, 0);
        chunk.emit_call(import, 1, 0);
        let store = chunk.code.len();
        chunk.emit_op_u16(Op::LOCAL_SET, 0, 0);
        let ret = chunk.code.len();
        chunk.emit_op(Op::RETURN, 0);
        let buckets = chunk.code.len().div_ceil(4);
        let mut ops = vec![Op::NOP; buckets];
        ops[0] = Op::CALL;
        ops[store / 4] = Op::LOCAL_SET;
        ops[ret / 4] = Op::RETURN;
        vm.chunks.push(chunk);
        vm.decoded_ops.push(ops);
        vm.decoded_operands.push(vec![0; buckets]);
        vm.native_call_sites.push(Vec::new());
        vm.local_fast_paths.push(Vec::new());
        vm.native_result_fast_paths.push(Vec::new());
        let target = ResolvedCallTarget::JsString(JsStringBuiltin::Length);
        if fallback {
            vm.chunk_call_target_cache.push(vec![Some(None)]);
            vm.resolved_import_table.push(target);
        } else {
            vm.chunk_call_target_cache.push(vec![Some(Some(target))]);
        }
        vm
    }

    #[test]
    fn unused_body_has_no_dense_native_row_then_entry_preserves_fusions() {
        for fallback in [false, true] {
            let mut vm = fixture(fallback);
            vm.prefill_native_call_sites(vec![(0, 0, 0)]).unwrap();
            assert!(vm.native_call_sites[0].is_empty());
            assert_eq!(vm.pending_native_call_sites[&0].len(), 1);
            // Nested-module completion discards transient lookup caches and
            // restores another module's import table. Deferred targets must
            // retain their original module identity across that boundary.
            vm.chunk_import_cache.clear();
            vm.chunk_call_target_cache.clear();
            vm.resolved_import_table = vec![ResolvedCallTarget::JsBoolean(
                crate::vm::JsBooleanBuiltin::Cast)];
            let mut eager = vec![None; vm.chunks[0].code.len().div_ceil(4)];
            eager[0] = Some(NativeCallSiteTarget::JsString(JsStringBuiltin::Length));
            let expected = VM::build_execution_fast_paths_for_chunk(&vm.chunks[0], Some(&eager));
            vm.ensure_execution_fast_paths(0);
            assert!(!vm.pending_native_call_sites.contains_key(&0));
            assert_eq!(vm.native_call_sites[0], eager);
            assert_eq!(vm.local_fast_paths[0], expected.0);
            assert_eq!(vm.native_result_fast_paths[0], expected.1);
            let sites = vm.native_call_sites[0].as_ptr();
            let locals = vm.local_fast_paths[0].as_ptr();
            vm.ensure_execution_fast_paths(0);
            assert_eq!(vm.native_call_sites[0].as_ptr(), sites);
            assert_eq!(vm.local_fast_paths[0].as_ptr(), locals);
        }
    }
}
