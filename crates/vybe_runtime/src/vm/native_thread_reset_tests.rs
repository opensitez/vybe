use super::*;
use std::sync::mpsc::{self, Sender};
use std::time::{Duration, Instant};

const WAIT: Duration = Duration::from_secs(5);

struct BlockedThread {
    release: Option<Sender<()>>,
    handle: Arc<NativeThread>,
}

impl BlockedThread {
    fn new(family: &NativeThreadFamily) -> Self {
        let (release, receive) = mpsc::channel();
        let handle = family.register(std::thread::spawn(move || {
            let _ = receive.recv_timeout(WAIT);
            vec![7]
        }));
        Self {
            release: Some(release),
            handle,
        }
    }

    fn finish(&mut self) {
        if let Some(release) = self.release.take() {
            let _ = release.send(());
        }
        assert_eq!(self.handle.join(), Ok(vec![7]));
    }
}

impl Drop for BlockedThread {
    fn drop(&mut self) {
        if let Some(release) = self.release.take() {
            let _ = release.send(());
        }
        let _ = self.handle.join();
    }
}

#[test]
fn active_native_child_refuses_reset_before_heap_roots_and_memory_change() {
    crate::heap::enable_tracking();
    let mut vm = VM::new();
    vm.memory = SharedMemory::from_vec(vec![1; 16]);
    let baseline = vm.snapshot();
    let object = crate::heap::alloc(Object::new());
    object
        .lock()
        .unwrap()
        .properties
        .insert("tenant".into(), Value::I32(9));
    vm.set_global_owned("tenant", Value::Object(object.clone()));
    vm.memory.atomic_store_i32(0, 42);
    vm.stack.push(Value::I32(99));
    let mut child = BlockedThread::new(&vm.native_threads);
    vm.thread_handles.insert(1, child.handle.clone());

    let started = Instant::now();
    assert!(matches!(
        vm.try_reset_to(&baseline),
        Err(ResetError::NativeThreadsActive { count: 1 })
    ));
    assert!(started.elapsed() < Duration::from_secs(1));
    assert_eq!(vm.memory.atomic_load_i32(0), 42);
    assert!(vm.has_global("tenant"));
    assert_eq!(
        object.lock().unwrap().properties.get("tenant"),
        Some(&Value::I32(9))
    );
    assert_eq!(vm.stack, vec![Value::I32(99)]);
    assert_eq!(vm.thread_handles.len(), 1);

    child.finish();
    vm.try_reset_to(&baseline).unwrap();
    assert!(!vm.has_global("tenant"));
    assert!(object.lock().unwrap().properties.is_empty());
    assert_eq!(vm.memory.atomic_load_i32(0), i32::from_le_bytes([1; 4]));
    assert!(vm.thread_handles.is_empty());
}

#[test]
fn reset_wrapper_panics_without_mutation_or_live_detachment() {
    let mut vm = VM::new();
    let baseline = vm.snapshot();
    vm.stack.push(Value::I32(31));
    let mut child = BlockedThread::new(&vm.native_threads);
    vm.thread_handles.insert(1, child.handle.clone());
    let refused = std::panic::catch_unwind(std::panic::AssertUnwindSafe(|| vm.reset_to(&baseline)));
    assert!(refused.is_err());
    assert_eq!(vm.stack, vec![Value::I32(31)]);
    assert_eq!(vm.thread_handles.len(), 1);
    child.finish();
    vm.reset_to(&baseline);
}

#[test]
fn exited_parent_cannot_hide_unfinished_grandchild_from_root() {
    let mut vm = VM::new();
    let baseline = vm.snapshot();
    let family = vm.native_threads.clone();
    let (send, receive) = mpsc::channel();
    let parent = vm.native_threads.register(std::thread::spawn(move || {
        let mut child_vm = VM::new();
        child_vm.native_threads = family;
        let child = BlockedThread::new(&child_vm.native_threads);
        child_vm.thread_handles.insert(1, child.handle.clone());
        send.send(child).unwrap();
        drop(child_vm);
        vec![0]
    }));
    let mut grandchild = receive.recv_timeout(WAIT).unwrap();
    assert_eq!(parent.join(), Ok(vec![0]));
    assert!(vm.thread_handles.is_empty());
    assert!(matches!(
        vm.ensure_threads_finished(),
        Err(ResetError::NativeThreadsActive { count: 1 })
    ));
    assert!(vm.try_reset_to(&baseline).is_err());
    grandchild.finish();
    vm.ensure_threads_finished().unwrap();
    vm.try_reset_to(&baseline).unwrap();
    assert!(vm.native_threads.handles.lock().unwrap().is_empty());
}

#[test]
fn finished_native_handles_are_joined_and_guest_outcomes_remain_available() {
    let mut vm = VM::new();
    let baseline = vm.snapshot();
    let mut child = BlockedThread::new(&vm.native_threads);
    child.release.take().unwrap().send(()).unwrap();
    let deadline = Instant::now() + WAIT;
    while !child
        .handle
        .state
        .lock()
        .unwrap()
        .handle
        .as_ref()
        .unwrap()
        .is_finished()
    {
        assert!(Instant::now() < deadline, "native child failed to finish");
        std::thread::sleep(Duration::from_millis(1));
    }
    vm.ensure_threads_finished().unwrap();
    assert!(child.handle.state.lock().unwrap().handle.is_none());
    let panic_child = vm
        .native_threads
        .register(std::thread::spawn(|| -> Vec<u8> {
            panic!("synthetic native child panic")
        }));
    assert!(panic_child.join().is_err());
    vm.ensure_threads_finished().unwrap();
    assert_eq!(child.handle.join(), Ok(vec![7]));
    assert!(panic_child.join().is_err());
    vm.try_reset_to(&baseline).unwrap();
}

#[test]
fn concurrent_guest_join_does_not_block_reset_refusal() {
    let mut vm = VM::new();
    let mut child = BlockedThread::new(&vm.native_threads);
    let joining = child.handle.state.lock().unwrap();
    let started = Instant::now();
    assert!(vm.ensure_threads_finished().is_err());
    assert!(started.elapsed() < Duration::from_secs(1));
    drop(joining);
    child.finish();
    vm.ensure_threads_finished().unwrap();
}

#[test]
fn reset_releases_oversized_stack_and_chunk_capacities() {
    let mut vm = VM::new();
    let baseline = vm.snapshot();
    const PEAK_BYTES: usize = 1024 * 1024;
    vm.stack
        .reserve_exact(PEAK_BYTES / std::mem::size_of::<Value>());
    vm.chunks
        .reserve_exact(PEAK_BYTES / std::mem::size_of::<Chunk>());
    assert!(vm.stack.capacity() * std::mem::size_of::<Value>() >= PEAK_BYTES / 2);
    assert!(vm.chunks.capacity() * std::mem::size_of::<Chunk>() >= PEAK_BYTES / 2);
    vm.stack.push(Value::I32(1));
    vm.chunks.push(Chunk::new("tenant"));
    vm.try_reset_to(&baseline).unwrap();
    assert!(vm.stack.is_empty());
    assert_eq!(vm.chunks.len(), baseline.chunks_len);
    assert!(vm.stack.capacity() * std::mem::size_of::<Value>() <= 64 * 1024);
    assert!(vm.chunks.capacity() * std::mem::size_of::<Chunk>() <= 64 * 1024);
}

#[test]
fn reset_restores_memory_limits_and_multi_memory_growth_metadata() {
    let mut vm = VM::new();
    vm.memory = SharedMemory::new(65536);
    vm.memory.set_max_pages(Some(2));
    let mut extra = SharedMemory::new(65536);
    extra.set_max_pages(Some(3));
    vm.extra_memories.push(extra);
    vm.extra_memory_max_pages = vec![Some(3)];
    let baseline = vm.snapshot();

    vm.memory.set_max_pages(Some(1));
    vm.extra_memories[0].set_max_pages(Some(1));
    vm.extra_memory_max_pages = vec![Some(1), Some(9)];
    vm.extra_memories.push(SharedMemory::new(65536));
    vm.try_reset_to(&baseline).unwrap();

    assert_eq!(vm.extra_memories.len(), 1);
    assert_eq!(vm.extra_memory_max_pages, vec![Some(3)]);
    assert_eq!(vm.mem_grow(0, 1), 1);
    assert_eq!(vm.mem_grow(0, 1), usize::MAX);
    assert_eq!(vm.mem_grow(1, 2), 1);
    assert_eq!(vm.mem_grow(1, 1), usize::MAX);
}
