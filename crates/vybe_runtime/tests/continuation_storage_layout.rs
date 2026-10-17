use std::{mem::size_of, sync::Mutex};
use vybe_runtime::{VM, Value, value::{ContinuationPhase, ContinuationState, Object, ObjectKind}};

#[test]
fn continuation_payload_does_not_inflate_common_object_headers() {
    assert!(size_of::<ContinuationState>() <= 80);
    assert!(size_of::<ObjectKind>() <= 128);
    assert!(size_of::<Object>() <= 256);
}

#[test]
fn cloning_continuation_keeps_entry_and_creates_independent_ready_storage() {
    let mut vm = VM::new();
    let original = ContinuationState {
        entry: Value::I32(17),
        saved: Box::new(Mutex::new(Some(vm.save_fiber()))),
        state: Mutex::new(ContinuationPhase::Suspended),
    };
    let cloned = original.clone();
    assert!(matches!(cloned.entry, Value::I32(17)));
    assert!(original.saved.lock().unwrap().is_some());
    assert!(cloned.saved.lock().unwrap().is_none());
    assert_eq!(*cloned.state.lock().unwrap(), ContinuationPhase::Ready);
    assert!(!std::ptr::eq(&*original.saved, &*cloned.saved));
    *cloned.state.lock().unwrap() = ContinuationPhase::Done;
    assert_eq!(*original.state.lock().unwrap(), ContinuationPhase::Suspended);
}
