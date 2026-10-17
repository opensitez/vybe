use vybe_runtime::{Chunk, Op, VM, Value};
use vybe_runtime::shared_provider::RuntimeHelperReceipt;
use vybe_runtime::value::{Function, Object, ObjectKind, Upvalue, UpvalueLocation};
use std::sync::{Arc, Mutex};

fn function(index: usize, arity: u8) -> Value {
    Value::Object(vybe_runtime::heap::alloc(Object {
        properties: Default::default(), type_id: 0, descriptor: None, fields: Vec::new(),
        kind: ObjectKind::Function(Function { name: None, arity, chunk_index: index, upvalues: Vec::new() }),
    }))
}

fn fixture() -> (VM, RuntimeHelperReceipt, Value) {
    let receipt = RuntimeHelperReceipt::new("__provider_test", "recipe-v1", 2, 1);
    let mut chunk = Chunk::new("name_is_not_authority");
    chunk.arity = 2;
    chunk.local_count = 2;
    chunk.runtime_helper_receipt = Some(receipt.clone());
    chunk.emit_op_u16(Op::LOCAL_GET, 0, 0);
    chunk.emit_op(Op::RETURN, 0);
    let value = function(0, 2);
    let mut vm = VM::new();
    vm.chunks.push(chunk);
    vm.set_global(receipt.global_name(), value.clone());
    (vm, receipt, value)
}

#[test]
fn requires_exact_process_identity_not_equal_labels() {
    let (vm, receipt, _) = fixture();
    let snapshot = vm.snapshot_runtime_helper(&receipt).unwrap();
    assert_eq!(snapshot.global_name(), "__provider_test");
    assert_eq!(snapshot.chunk_index(), 0);
    assert!(snapshot.has_receipt(&receipt));
    assert!(vm.validate_runtime_helper(&snapshot));
    let impostor = RuntimeHelperReceipt::new("__provider_test", "recipe-v1", 2, 1);
    assert!(!snapshot.has_receipt(&impostor));
    assert!(vm.snapshot_runtime_helper(&impostor).is_none());
}

#[test]
fn guest_names_and_unstamped_chunks_cannot_attest() {
    let (mut vm, receipt, _) = fixture();
    vm.chunks[0].name = "__provider_test".into();
    vm.chunks[0].runtime_helper_receipt = None;
    assert!(vm.snapshot_runtime_helper(&receipt).is_none());
    assert!(Chunk::new("__provider_test").runtime_helper_receipt.is_none());
}

#[test]
fn rebinding_and_mutated_callable_state_invalidate_snapshots() {
    let (mut vm, receipt, original) = fixture();
    let snapshot = vm.snapshot_runtime_helper(&receipt).unwrap();
    vm.set_global(receipt.global_name(), function(0, 2));
    assert!(!vm.validate_runtime_helper(&snapshot));
    vm.set_global(receipt.global_name(), original.clone());
    assert!(vm.validate_runtime_helper(&snapshot));
    let Value::Object(object) = original else { unreachable!() };
    let mut locked = object.lock().unwrap();
    let ObjectKind::Function(function) = &mut locked.kind else { unreachable!() };
    function.upvalues.push(Arc::new(Mutex::new(Upvalue { location: UpvalueLocation::Closed(Value::Null) })));
    drop(locked);
    assert!(!vm.validate_runtime_helper(&snapshot));
}

#[test]
fn abi_mismatches_are_not_reusable() {
    let (mut vm, receipt, _) = fixture();
    let original = vm.chunks[0].clone();
    for variant in 0..6 {
        let mut chunk = original.clone();
        match variant {
            0 => chunk.takes_receiver = true,
            1 => chunk.capture_count = 1,
            2 => chunk.result_arity = 2,
            3 => chunk.direct_entry_has_captures = Some(true),
            4 => chunk.is_generator = true,
            _ => chunk.is_async = true,
        }
        vm.chunks[0] = chunk;
        assert!(vm.snapshot_runtime_helper(&receipt).is_none());
    }
}

#[test]
fn snapshots_are_vm_local_even_when_bindings_and_chunks_are_shared() {
    let (vm, receipt, binding) = fixture();
    let snapshot = vm.snapshot_runtime_helper(&receipt).unwrap();
    let mut other = VM::new();
    other.chunks = vm.chunks.clone();
    other.set_global(receipt.global_name(), binding);
    assert!(other.snapshot_runtime_helper(&receipt).is_some());
    assert!(!other.validate_runtime_helper(&snapshot));
}

#[test]
fn unassigned_global_slots_do_not_attest_stale_values() {
    let (mut vm, receipt, _) = fixture();
    let snapshot = vm.snapshot_runtime_helper(&receipt).unwrap();
    let slot = vm.global_index[receipt.global_name()] as usize;
    vm.globals_assigned[slot] = false;
    assert!(vm.snapshot_runtime_helper(&receipt).is_none());
    assert!(!vm.validate_runtime_helper(&snapshot));
    vm.globals_assigned[slot] = true;
    assert!(vm.validate_runtime_helper(&snapshot));
}

#[test]
fn attached_step_debugger_keeps_global_initialization_observable() {
    let (mut vm, receipt, _) = fixture();
    let snapshot = vm.snapshot_runtime_helper(&receipt).unwrap();
    let (_commands, command_rx) = std::sync::mpsc::channel();
    let (event_tx, _events) = std::sync::mpsc::channel();
    vm.attach_debugger(command_rx, event_tx, false);
    assert!(vm.snapshot_runtime_helper(&receipt).is_none());
    assert!(!vm.validate_runtime_helper(&snapshot));
    vm.detach_debugger();
    assert!(vm.validate_runtime_helper(&snapshot));
}

#[test]
fn snapshots_do_not_keep_rebound_functions_alive() {
    let (mut vm, receipt, original) = fixture();
    let Value::Object(object) = &original else { unreachable!() };
    let strong_count = Arc::strong_count(object);
    let weak = Arc::downgrade(object);
    let snapshot = vm.snapshot_runtime_helper(&receipt).unwrap();
    assert_eq!(Arc::strong_count(object), strong_count);
    vm.set_global(receipt.global_name(), Value::Null);
    drop(original);
    assert!(weak.upgrade().is_none());
    assert!(!vm.validate_runtime_helper(&snapshot));
}
