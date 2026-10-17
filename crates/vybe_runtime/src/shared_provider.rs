//! Admission for compiler-owned runtime providers across compilation units.
//!
//! These are host compilation/linkage receipts, never guest instructions,
//! imported functions, language tags, or serialized Wasm metadata. Trusted
//! compiler recipes mint one identity per exact implementation/ABI contract.

use std::sync::{Arc, Mutex, Weak};
use crate::{Value, VM};
use crate::value::{Object, ObjectKind};

#[derive(Clone, Debug)]
pub struct RuntimeHelperReceipt(Arc<Identity>);

#[derive(Debug)]
struct Identity {
    global_name: &'static str,
    version: &'static str,
    arity: u8,
    result_arity: u8,
}

impl RuntimeHelperReceipt {
    /// Mint from trusted Rust compilation code, normally in a recipe's
    /// OnceLock. Identical labels do not grant identical authority. Do not
    /// reconstruct this from a guest name, body hash or custom debug section.
    pub fn new(global_name: &'static str, version: &'static str, arity: u8, result_arity: u8) -> Self {
        Self(Arc::new(Identity { global_name, version, arity, result_arity }))
    }

    pub fn global_name(&self) -> &'static str { self.0.global_name }
    pub fn version(&self) -> &'static str { self.0.version }
    pub fn arity(&self) -> u8 { self.0.arity }
    pub fn result_arity(&self) -> u8 { self.0.result_arity }
    fn same_identity(&self, other: &Self) -> bool { Arc::ptr_eq(&self.0, &other.0) }
}

#[derive(Clone, Debug)]
pub struct InstalledRuntimeProvider {
    scope: Arc<()>,
    receipt: RuntimeHelperReceipt,
    binding: Weak<Mutex<Object>>,
    chunk_index: usize,
}

impl InstalledRuntimeProvider {
    pub fn global_name(&self) -> &'static str { self.receipt.global_name() }
    pub fn chunk_index(&self) -> usize { self.chunk_index }
    pub fn has_receipt(&self, receipt: &RuntimeHelperReceipt) -> bool {
        self.receipt.same_identity(receipt)
    }
}

impl VM {
    fn runtime_helper_target(&self, receipt: &RuntimeHelperReceipt) -> Option<(usize, &Value)> {
        // Global initialization has debugger-visible before/after events.
        // Reuse must not suppress those events under an attached debugger.
        if self.debugger.is_some() { return None; }
        let slot = *self.global_index.get(receipt.global_name())? as usize;
        if !self.globals_assigned.get(slot).copied().unwrap_or(false) { return None; }
        let binding = self.globals.get(slot)?;
        let Value::Object(object) = binding else { return None; };
        let object = object.lock().ok()?;
        let ObjectKind::Function(function) = &object.kind else { return None; };
        if function.arity != receipt.arity() || !function.upvalues.is_empty()
            || object.properties.contains_key("__bound_args")
        { return None; }
        let chunk = self.chunks.get(function.chunk_index)?;
        if !chunk.runtime_helper_receipt.as_ref()?.same_identity(receipt)
            || chunk.arity != receipt.arity() || chunk.result_arity != receipt.result_arity()
            || chunk.local_count < u16::from(chunk.arity)
            || chunk.takes_receiver || chunk.capture_count != 0
            || chunk.direct_entry_has_captures == Some(true)
            || chunk.is_async || chunk.is_generator || chunk.func_sig.is_some()
        { return None; }
        Some((function.chunk_index, binding))
    }

    /// Observe a live, capture-free provider with this exact recipe identity.
    /// The caller must revalidate the snapshot before publishing code which
    /// omitted a local implementation. Stale snapshots are not a cache hit.
    pub fn snapshot_runtime_helper(&self, receipt: &RuntimeHelperReceipt) -> Option<InstalledRuntimeProvider> {
        let (chunk_index, binding) = self.runtime_helper_target(receipt)?;
        let Value::Object(binding) = binding else { return None; };
        Some(InstalledRuntimeProvider {
            scope: self.runtime_helper_scope.clone(), receipt: receipt.clone(),
            binding: Arc::downgrade(binding), chunk_index,
        })
    }

    pub fn validate_runtime_helper(&self, snapshot: &InstalledRuntimeProvider) -> bool {
        if !Arc::ptr_eq(&self.runtime_helper_scope, &snapshot.scope) { return false; }
        let Some((index, binding)) = self.runtime_helper_target(&snapshot.receipt) else { return false; };
        index == snapshot.chunk_index && match binding {
            Value::Object(current) => Weak::ptr_eq(&Arc::downgrade(current), &snapshot.binding),
            _ => false,
        }
    }

    /// Read an installed provider's import declaration in its owning module.
    /// Linked function bodies normally have an empty local import table: their
    /// CALL indices refer to the module's first chunk. Match dispatch's own-
    /// table-before-owner lookup without touching resolved import caches.
    /// A missing declaration remains unknown; the transient VM-wide resolved
    /// table cannot supply a stable declaration for compiler diagnostics.
    pub fn runtime_helper_import(
        &self,
        provider: &InstalledRuntimeProvider,
        index: usize,
    ) -> Option<&crate::chunk::Import> {
        if !self.validate_runtime_helper(provider) { return None; }
        let chunk_index = provider.chunk_index;
        let owner = self.chunk_import_owner.get(chunk_index).copied().unwrap_or(chunk_index);
        self.chunks.get(chunk_index)?.imports.get(index).or_else(|| {
            self.chunks.get(owner)?.imports.get(index)
        })
    }
}

#[cfg(test)]
mod provider_import_tests {
    use super::*;
    use crate::chunk::Chunk;
    use crate::value::Function;

    #[test]
    fn provider_import_uses_own_then_bound_owner_and_rejects_stale_authority() {
        let receipt = RuntimeHelperReceipt::new("__provider_import_test", "v1", 1, 1);
        let mut helper = Chunk::new("helper");
        helper.arity = 1;
        helper.local_count = 1;
        helper.result_arity = 1;
        helper.runtime_helper_receipt = Some(receipt.clone());
        let mut owner = Chunk::new("owner");
        owner.imports.intern("owner", "first");
        owner.imports.intern("owner", "second");
        let mut vm = VM::new();
        vm.chunks = vec![owner, helper];
        vm.chunk_import_owner = vec![0, 0];
        let mut object = Object::new();
        object.kind = ObjectKind::Function(Function {
            name: None, arity: 1, chunk_index: 1, upvalues: Vec::new(),
        });
        vm.set_global(receipt.global_name(), Value::Object(crate::heap::alloc(object)));
        let provider = vm.snapshot_runtime_helper(&receipt).unwrap();
        assert_eq!(vm.runtime_helper_import(&provider, 0).unwrap().module, "owner");
        vm.chunks[1].imports.intern("own", "first");
        assert_eq!(vm.runtime_helper_import(&provider, 0).unwrap().module, "own");
        assert_eq!(vm.runtime_helper_import(&provider, 1).unwrap().name, "second");
        assert!(vm.runtime_helper_import(&provider, 2).is_none());
        let foreign = VM::new();
        assert!(foreign.runtime_helper_import(&provider, 0).is_none());
        vm.set_global(receipt.global_name(), Value::Null);
        assert!(vm.runtime_helper_import(&provider, 0).is_none());
    }
}
