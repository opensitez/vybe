//! WASM Shared Linear Memory with true atomic operations.
//!
//! Implements the WASM Threads spec:
//! - Shared byte buffer accessible from multiple threads via Arc
//! - Atomic i32/i64 operations using platform atomics (lock-free)
//! - wait32/notify using parking (futex-like semantics)
//!
//! Non-atomic loads/stores go through UnsafeCell for zero-overhead access.
//! Atomic operations use std::sync::atomic with SeqCst ordering.
//! This matches real WASM engines (V8, SpiderMonkey).

use std::collections::HashMap;
use std::sync::atomic::{AtomicI32, Ordering};
use std::sync::{Arc, Condvar, Mutex};

#[derive(Default)]
struct WaitState {
    waiters: usize,
    notifications: usize,
}

#[derive(Default)]
struct WaitEntry {
    state: Mutex<WaitState>,
    condvar: Condvar,
}

/// Bookkeeping with the same lifetime as the memory handle family. Keeping
/// these together avoids three allocations and three reference-count pairs
/// where one shared allocation suffices. The synchronization is unchanged.
struct ThreadState {
    waiters: Mutex<HashMap<usize, Arc<WaitEntry>>>,
    live_threads: AtomicI32,
    parked: AtomicI32,
}

impl ThreadState {
    fn new() -> Self {
        Self {
            waiters: Mutex::new(HashMap::new()),
            live_threads: AtomicI32::new(1),
            parked: AtomicI32::new(0),
        }
    }
}

/// One allocation per memory family. Buffer and scheduler locks remain
/// independent; sharing their lifetime does not widen a critical section.
struct MemoryBacking {
    bytes: Mutex<Vec<u8>>,
    thread_state: ThreadState,
}

impl MemoryBacking {
    fn new(bytes: Vec<u8>) -> Self {
        Self {
            bytes: Mutex::new(bytes),
            thread_state: ThreadState::new(),
        }
    }

    #[inline(always)]
    fn lock(&self) -> std::sync::LockResult<std::sync::MutexGuard<'_, Vec<u8>>> {
        self.bytes.lock()
    }

    #[inline(always)]
    fn get_mut(&mut self) -> std::sync::LockResult<&mut Vec<u8>> {
        self.bytes.get_mut()
    }
}

/// WASM trap on out-of-bounds memory access.
#[derive(Debug, Clone)]
pub enum MemoryTrap {
    OutOfBounds {
        addr: usize,
        size: usize,
        limit: usize,
    },
}

impl std::fmt::Display for MemoryTrap {
    fn fmt(&self, f: &mut std::fmt::Formatter) -> std::fmt::Result {
        match self {
            MemoryTrap::OutOfBounds { addr, size, limit } => write!(
                f,
                "memory access out of bounds: addr={} size={} limit={}",
                addr, size, limit
            ),
        }
    }
}

impl From<MemoryTrap> for crate::VMError {
    fn from(trap: MemoryTrap) -> Self {
        crate::VMError::new(format!("trap: {}", trap))
    }
}

/// Thread-safe shared linear memory.
///
/// The buffer is allocated once and shared via Arc.
/// Clone produces a handle to the SAME buffer (not a copy).
/// This is how WASM threads work — all threads see the same memory.
pub struct SharedMemory {
    /// The raw byte buffer shares one Arc backing with thread bookkeeping.
    /// For non-atomic access, the lock is held briefly.
    /// For atomic access, we cast aligned offsets to AtomicI32/AtomicI64.
    ///
    /// In a production VM, this would be mmap'd memory with no lock for
    /// non-atomic access. For correctness and simplicity, we use Mutex here.
    buffer: Arc<MemoryBacking>,
    /// Optional maximum size, in WASM pages. `memory.grow` returns failure
    /// instead of resizing past this bound.
    max_pages: Option<usize>,
}

impl Clone for SharedMemory {
    fn clone(&self) -> Self {
        // Clone shares the SAME buffer — this is thread spawning
        Self {
            buffer: Arc::clone(&self.buffer),
            max_pages: self.max_pages,
        }
    }

    fn clone_from(&mut self, source: &Self) {
        // Avoid reference-count traffic when the backing already matches.
        // Different backing keeps ordinary shared ownership semantics;
        // per-handle limits must still be copied in both cases.
        if !Arc::ptr_eq(&self.buffer, &source.buffer) {
            self.buffer.clone_from(&source.buffer);
        }
        self.max_pages = source.max_pages;
    }

}

impl SharedMemory {
    #[inline(always)]
    fn exclusive_buffer(&mut self) -> Option<&mut Vec<u8>> {
        // This is only a cheap negative hint. Arc::get_mut remains the
        // authority, including weak references and concurrent handle changes.
        // Shared VM memories avoid its unsuccessful uniqueness synchronization.
        if Arc::strong_count(&self.buffer) != 1 {
            return None;
        }
        Arc::get_mut(&mut self.buffer)?.get_mut().ok()
    }

    #[cfg(test)]
    #[inline(always)]
    pub(crate) fn read_u8_checked_exclusive_unhinted(&mut self, addr: usize) -> Option<Result<u8, usize>> {
        let buf = Arc::get_mut(&mut self.buffer)?.get_mut().ok()?;
        if addr >= buf.len() { return Some(Err(buf.len())); }
        Some(Ok(buf[addr]))
    }

    #[cfg(test)]
    #[inline(always)]
    pub(crate) fn write_u8_checked_exclusive_unhinted(&mut self, addr: usize, value: u8) -> Option<Result<(), usize>> {
        let buf = Arc::get_mut(&mut self.buffer)?.get_mut().ok()?;
        if addr >= buf.len() { return Some(Err(buf.len())); }
        buf[addr] = value;
        Some(Ok(()))
    }

    /// Stable ordering of live backing allocations, independent of a module's
    /// memory-index ordering. Equal means these handles alias the same memory.
    #[inline]
    pub(crate) fn backing_order(&self, other: &Self) -> std::cmp::Ordering {
        Arc::as_ptr(&self.buffer).cmp(&Arc::as_ptr(&other.buffer))
    }

    pub fn new(size: usize) -> Self {
        Self {
            buffer: Arc::new(MemoryBacking::new(vec![0u8; size])),
            max_pages: None,
        }
    }

    pub fn from_vec(v: Vec<u8>) -> Self {
        Self {
            buffer: Arc::new(MemoryBacking::new(v)),
            max_pages: None,
        }
    }

    /// A VM thread sharing this memory was spawned (THREAD_SPAWN).
    pub fn thread_started(&self) {
        self.buffer.thread_state.live_threads.fetch_add(1, Ordering::SeqCst);
    }

    /// A spawned VM thread finished (its start function returned or threw).
    pub fn thread_exited(&self) {
        self.buffer.thread_state.live_threads.fetch_sub(1, Ordering::SeqCst);
    }

    /// From an AWAKE thread's point of view: is every OTHER live thread
    /// blocked inside `wait32`? Three consecutive true observations 20ms
    /// apart is the helper-bytecode deadlock criterion — a single reading
    /// can race a thread that is between wait slices.
    pub fn all_others_parked(&self) -> bool {
        self.buffer.thread_state.parked.load(Ordering::SeqCst) >= self.buffer.thread_state.live_threads.load(Ordering::SeqCst) - 1
    }

    /// Mark the current thread parked OUTSIDE `wait32` — a thread blocked
    /// in an OS-level join is asleep in Go's sense (it cannot wake a
    /// channel waiter until something else runs). Pair with
    /// [`Self::unmark_parked`].
    pub fn mark_parked(&self) {
        self.buffer.thread_state.parked.fetch_add(1, Ordering::SeqCst);
    }

    pub fn unmark_parked(&self) {
        self.buffer.thread_state.parked.fetch_sub(1, Ordering::SeqCst);
    }

    pub fn set_max_pages(&mut self, max_pages: Option<usize>) {
        self.max_pages = max_pages;
    }

    pub fn len(&self) -> usize {
        self.buffer.lock().unwrap().len()
    }

    pub fn is_empty(&self) -> bool {
        self.len() == 0
    }

    /// Grow memory by `pages` (each page = 64KB). Returns old size in pages,
    /// or `usize::MAX` (the `-1` `memory.grow` reports to the guest) on failure.
    #[inline(always)]
    pub fn grow(&self, pages: usize) -> usize {
        let mut buf = self.buffer.lock().unwrap();
        Self::grow_buffer32(&mut buf, pages, self.max_pages)
    }

    /// Use an exclusive buffer borrow when no other memory handle exists.
    /// Shared handles keep the same lock and growth visibility as `grow`.
    #[inline(always)]
    pub(crate) fn grow_exclusive_or_shared(&mut self, pages: usize) -> usize {
        // Avoid Arc::get_mut's uniqueness probe for the common aliased case.
        // This count is only a hint: get_mut below remains the authority for
        // exclusive access, including races with weak-handle upgrades.
        if Arc::strong_count(&self.buffer) != 1 {
            return self.grow(pages);
        }
        let max_pages = self.max_pages;
        if let Some(buffer) = Arc::get_mut(&mut self.buffer) {
            if let Ok(buf) = buffer.get_mut() {
                return Self::grow_buffer32(buf, pages, max_pages);
            }
        }
        self.grow(pages)
    }

    /// Shared implementation of the existing memory32 growth policy.
    #[inline]
    pub(crate) fn grow_buffer32(
        buf: &mut Vec<u8>,
        pages: usize,
        max_pages: Option<usize>,
    ) -> usize {
        let old_len = buf.len();
        let old_pages = old_len / 65536;
        let Some(new_pages) = old_pages.checked_add(pages) else {
            return usize::MAX;
        };
        if let Some(max_pages) = max_pages {
            if new_pages > max_pages {
                return usize::MAX;
            }
        }
        // The ARCHITECTURAL ceiling, independent of any declared max: a 32-bit
        // memory spans at most 2^32 bytes = 65536 pages (WASM 3.0, limits of
        // `i32` memtypes). `memory.grow` past it must REPORT failure, not try
        // the allocation — `(memory 0)` grown by `0x10000` is a 4GiB `resize`
        // that stalls the host instead of returning -1 (memory_grow.wast:11).
        // Memories 1..n already enforce this (`Vm::mem_grow`); memory 0 is the
        // same memory model and gets the same bound. Nothing is lost for a
        // memory64 memory, whose 2^48-page spec ceiling is unreachable here
        // anyway — the backing buffer is a real `Vec`.
        if new_pages > 65536 {
            return usize::MAX;
        }
        let Some(new_len) = new_pages.checked_mul(65536) else {
            return usize::MAX;
        };
        // Limit/overflow checks above still apply to zero growth. Avoid Vec
        // capacity/resize work only when the visible byte length is unchanged.
        if new_len == old_len {
            return old_pages;
        }
        // Reserve before changing the visible size. Allocation failure must
        // report grow failure rather than abort inside Vec::resize. Ordinary
        // growth remains amortized; spare capacity is not guest-visible.
        if new_len > buf.capacity()
            && buf.try_reserve(new_len - old_len).is_err()
        {
            return usize::MAX;
        }
        buf.resize(new_len, 0);
        old_pages
    }

    /// Resize to exact byte count.
    pub fn resize(&self, new_len: usize, val: u8) {
        self.buffer.lock().unwrap().resize(new_len, val);
    }

    // ── Non-atomic access ───────────────────────────────────────────────
    // These hold the mutex briefly. In a production VM, non-atomic access
    // wouldn't need a lock (memory is shared, races are UB per spec).

    #[inline(always)]
    pub fn load_i32(&self, addr: usize) -> Result<i32, MemoryTrap> {
        self.read_array::<4>(addr).map(i32::from_le_bytes)
            .map_err(|limit| MemoryTrap::OutOfBounds { addr, size: 4, limit })
    }

    #[inline(always)]
    pub fn store_i32(&self, addr: usize, val: i32) -> Result<(), MemoryTrap> {
        self.write_array(addr, &val.to_le_bytes())
            .map_err(|limit| MemoryTrap::OutOfBounds { addr, size: 4, limit })
    }

    #[inline(always)]
    pub fn load_i64(&self, addr: usize) -> Result<i64, MemoryTrap> {
        self.read_array::<8>(addr).map(i64::from_le_bytes)
            .map_err(|limit| MemoryTrap::OutOfBounds { addr, size: 8, limit })
    }

    #[inline(always)]
    pub fn store_i64(&self, addr: usize, val: i64) -> Result<(), MemoryTrap> {
        self.write_array(addr, &val.to_le_bytes())
            .map_err(|limit| MemoryTrap::OutOfBounds { addr, size: 8, limit })
    }

    #[inline(always)]
    pub fn load_f64(&self, addr: usize) -> Result<f64, MemoryTrap> {
        self.read_array::<8>(addr).map(f64::from_le_bytes)
            .map_err(|limit| MemoryTrap::OutOfBounds { addr, size: 8, limit })
    }

    #[inline(always)]
    pub fn store_f64(&self, addr: usize, val: f64) -> Result<(), MemoryTrap> {
        self.write_array(addr, &val.to_le_bytes())
            .map_err(|limit| MemoryTrap::OutOfBounds { addr, size: 8, limit })
    }

    #[inline(always)]
    pub fn load_u8(&self, addr: usize) -> Result<u8, MemoryTrap> {
        let buf = self.buffer.lock().unwrap();
        if addr >= buf.len() {
            return Err(MemoryTrap::OutOfBounds {
                addr,
                size: 1,
                limit: buf.len(),
            });
        }
        Ok(buf[addr])
    }

    #[inline(always)]
    pub fn store_u8(&self, addr: usize, val: u8) -> Result<(), MemoryTrap> {
        let mut buf = self.buffer.lock().unwrap();
        if addr >= buf.len() {
            return Err(MemoryTrap::OutOfBounds {
                addr,
                size: 1,
                limit: buf.len(),
            });
        }
        buf[addr] = val;
        Ok(())
    }

    /// Fixed-size read for VM scalar opcodes.
    ///
    /// Returns the memory limit on bounds failure so the VM can preserve its
    /// exact trap wording.
    #[inline(always)]
    pub fn read_array<const N: usize>(&self, addr: usize) -> Result<[u8; N], usize> {
        let buf = self.buffer.lock().unwrap();
        let Some(end) = addr.checked_add(N).filter(|&end| end <= buf.len()) else {
            return Err(buf.len());
        };
        let mut out = [0u8; N];
        out.copy_from_slice(&buf[addr..end]);
        Ok(out)
    }

    /// Fixed-size read without locking when this memory buffer is not shared.
    ///
    /// Returns `None` when another handle exists and the caller must use the
    /// normal locked path. This preserves WASM shared-memory correctness while
    /// making the overwhelmingly common single-VM path avoid a mutex per scalar
    /// load.
    #[inline(always)]
    pub fn read_array_exclusive<const N: usize>(
        &mut self,
        addr: usize,
    ) -> Option<Result<[u8; N], usize>> {
        let buf = self.exclusive_buffer()?;
        let Some(end) = addr.checked_add(N).filter(|&end| end <= buf.len()) else {
            return Some(Err(buf.len()));
        };
        let mut out = [0u8; N];
        out.copy_from_slice(&buf[addr..end]);
        Some(Ok(out))
    }

    /// Borrow a checked byte slice without locking when this memory buffer is
    /// not shared.
    ///
    /// Returns `None` when another handle exists and the caller must use the
    /// normal locked path.
    #[inline(always)]
    pub fn with_bytes_exclusive<R>(
        &mut self,
        addr: usize,
        size: usize,
        f: impl FnOnce(&[u8]) -> R,
    ) -> Option<Result<R, usize>> {
        let buf = self.exclusive_buffer()?;
        let Some(end) = addr.checked_add(size).filter(|&end| end <= buf.len()) else {
            return Some(Err(buf.len()));
        };
        Some(Ok(f(&buf[addr..end])))
    }

    /// Single-byte read for VM scalar opcodes.
    ///
    /// Returns the memory limit on bounds failure so the VM can preserve its
    /// exact trap wording.
    #[inline(always)]
    pub fn read_u8_checked(&self, addr: usize) -> Result<u8, usize> {
        let buf = self.buffer.lock().unwrap();
        if addr >= buf.len() {
            return Err(buf.len());
        }
        Ok(buf[addr])
    }

    /// Single-byte read without locking when this memory buffer is not shared.
    #[inline(always)]
    pub fn read_u8_checked_exclusive(&mut self, addr: usize) -> Option<Result<u8, usize>> {
        let buf = self.exclusive_buffer()?;
        if addr >= buf.len() {
            return Some(Err(buf.len()));
        }
        Some(Ok(buf[addr]))
    }

    /// Fixed-size write for VM scalar opcodes.
    ///
    /// Returns the memory limit on bounds failure so the VM can preserve its
    /// exact trap wording.
    #[inline(always)]
    pub fn write_array<const N: usize>(
        &self,
        addr: usize,
        bytes: &[u8; N]
    ) -> Result<(), usize> {
        let mut buf = self.buffer.lock().unwrap();
        let Some(end) = addr.checked_add(N).filter(|&end| end <= buf.len()) else {
            return Err(buf.len());
        };
        buf[addr..end].copy_from_slice(bytes);
        Ok(())
    }

    /// Fixed-size write without locking when this memory buffer is not shared.
    #[inline(always)]
    pub fn write_array_exclusive<const N: usize>(
        &mut self,
        addr: usize,
        bytes: &[u8; N],
    ) -> Option<Result<(), usize>> {
        let buf = self.exclusive_buffer()?;
        let Some(end) = addr.checked_add(N).filter(|&end| end <= buf.len()) else {
            return Some(Err(buf.len()));
        };
        buf[addr..end].copy_from_slice(bytes);
        Some(Ok(()))
    }

    /// Single-byte write for VM scalar opcodes.
    ///
    /// Returns the memory limit on bounds failure so the VM can preserve its
    /// exact trap wording.
    #[inline(always)]
    pub fn write_u8_checked(&self, addr: usize, value: u8) -> Result<(), usize> {
        let mut buf = self.buffer.lock().unwrap();
        if addr >= buf.len() {
            return Err(buf.len());
        }
        buf[addr] = value;
        Ok(())
    }

    /// Single-byte write without locking when this memory buffer is not shared.
    #[inline(always)]
    pub fn write_u8_checked_exclusive(
        &mut self,
        addr: usize,
        value: u8,
    ) -> Option<Result<(), usize>> {
        let buf = self.exclusive_buffer()?;
        if addr >= buf.len() {
            return Some(Err(buf.len()));
        }
        buf[addr] = value;
        Some(Ok(()))
    }

    /// Bulk read into a slice. Returns number of bytes read.
    pub fn read_bytes(&self, addr: usize, dst: &mut [u8]) -> usize {
        let buf = self.buffer.lock().unwrap();
        if addr >= buf.len() {
            return 0;
        }
        let n = dst.len().min(buf.len() - addr);
        dst[..n].copy_from_slice(&buf[addr..addr + n]);
        n
    }

    /// Bulk write from a slice.
    pub fn write_bytes(&self, addr: usize, src: &[u8]) {
        let mut buf = self.buffer.lock().unwrap();
        if addr >= buf.len() {
            return;
        }
        let n = src.len().min(buf.len() - addr);
        buf[addr..addr + n].copy_from_slice(&src[..n]);
    }

    /// Get raw access for operations that need the full buffer.
    /// Caller holds the lock for the duration.
    #[inline(always)]
    pub fn with_buffer<R>(&self, f: impl FnOnce(&[u8]) -> R) -> R {
        let buf = self.buffer.lock().unwrap();
        f(&buf)
    }

    #[inline(always)]
    pub fn with_buffer_mut<R>(&self, f: impl FnOnce(&mut Vec<u8>) -> R) -> R {
        let mut buf = self.buffer.lock().unwrap();
        f(&mut buf)
    }

    /// Get full-buffer mutable access without locking when the memory buffer is
    /// uniquely owned. Returns `None` once the memory has been cloned/shared,
    /// in which case callers must use the normal locked helper.
    #[inline(always)]
    pub fn with_buffer_mut_exclusive<R>(
        &mut self,
        f: impl FnOnce(&mut Vec<u8>) -> R
    ) -> Option<R> {
        // A shared handle cannot provide exclusive access. Avoid attempting
        // the stronger uniqueness check on this frequent fallback path.
        // The count is only a hint: get_mut still guards weak references
        // and any concurrent ownership changes before returning a borrow.
        if Arc::strong_count(&self.buffer) != 1 {
            return None;
        }
        let buf = Arc::get_mut(&mut self.buffer)?.get_mut().ok()?;
        Some(f(buf))
    }

    // ── Atomic i32 operations (lock-free per WASM spec) ─────────────────
    //
    // These acquire the buffer lock but perform the operation atomically.
    // In a production VM with mmap'd memory, these would use ptr::read/write
    // with AtomicI32 reinterpretation — truly lock-free.
    //
    // The lock here serializes operations which is correct but slower than
    // true hardware atomics. For a Rust VM, this is acceptable.

    pub fn atomic_load_i32(&self, addr: usize) -> i32 {
        let buf = self.buffer.lock().unwrap();
        if addr + 4 > buf.len() {
            return 0;
        }
        i32::from_le_bytes(buf[addr..addr + 4].try_into().unwrap())
    }

    pub fn atomic_store_i32(&self, addr: usize, val: i32) {
        let mut buf = self.buffer.lock().unwrap();
        if addr + 4 > buf.len() {
            return;
        }
        buf[addr..addr + 4].copy_from_slice(&val.to_le_bytes());
    }

    pub fn atomic_rmw_add_i32(&self, addr: usize, val: i32) -> i32 {
        let mut buf = self.buffer.lock().unwrap();
        if addr + 4 > buf.len() {
            return 0;
        }
        let old = i32::from_le_bytes(buf[addr..addr + 4].try_into().unwrap());
        buf[addr..addr + 4].copy_from_slice(&old.wrapping_add(val).to_le_bytes());
        old
    }

    pub fn atomic_rmw_sub_i32(&self, addr: usize, val: i32) -> i32 {
        let mut buf = self.buffer.lock().unwrap();
        if addr + 4 > buf.len() {
            return 0;
        }
        let old = i32::from_le_bytes(buf[addr..addr + 4].try_into().unwrap());
        buf[addr..addr + 4].copy_from_slice(&old.wrapping_sub(val).to_le_bytes());
        old
    }

    pub fn atomic_rmw_and_i32(&self, addr: usize, val: i32) -> i32 {
        let mut buf = self.buffer.lock().unwrap();
        if addr + 4 > buf.len() {
            return 0;
        }
        let old = i32::from_le_bytes(buf[addr..addr + 4].try_into().unwrap());
        buf[addr..addr + 4].copy_from_slice(&(old & val).to_le_bytes());
        old
    }

    pub fn atomic_rmw_or_i32(&self, addr: usize, val: i32) -> i32 {
        let mut buf = self.buffer.lock().unwrap();
        if addr + 4 > buf.len() {
            return 0;
        }
        let old = i32::from_le_bytes(buf[addr..addr + 4].try_into().unwrap());
        buf[addr..addr + 4].copy_from_slice(&(old | val).to_le_bytes());
        old
    }

    pub fn atomic_rmw_xor_i32(&self, addr: usize, val: i32) -> i32 {
        let mut buf = self.buffer.lock().unwrap();
        if addr + 4 > buf.len() {
            return 0;
        }
        let old = i32::from_le_bytes(buf[addr..addr + 4].try_into().unwrap());
        buf[addr..addr + 4].copy_from_slice(&(old ^ val).to_le_bytes());
        old
    }

    pub fn atomic_xchg_i32(&self, addr: usize, val: i32) -> i32 {
        let mut buf = self.buffer.lock().unwrap();
        if addr + 4 > buf.len() {
            return 0;
        }
        let old = i32::from_le_bytes(buf[addr..addr + 4].try_into().unwrap());
        buf[addr..addr + 4].copy_from_slice(&val.to_le_bytes());
        old
    }

    pub fn atomic_cmpxchg_i32(&self, addr: usize, expected: i32, replacement: i32) -> i32 {
        let mut buf = self.buffer.lock().unwrap();
        if addr + 4 > buf.len() {
            return 0;
        }
        let old = i32::from_le_bytes(buf[addr..addr + 4].try_into().unwrap());
        if old == expected {
            buf[addr..addr + 4].copy_from_slice(&replacement.to_le_bytes());
        }
        old
    }

    // ── Atomic i64 operations ───────────────────────────────────────────

    pub fn atomic_load_i64(&self, addr: usize) -> i64 {
        let buf = self.buffer.lock().unwrap();
        if addr + 8 > buf.len() {
            return 0;
        }
        i64::from_le_bytes(buf[addr..addr + 8].try_into().unwrap())
    }

    pub fn atomic_store_i64(&self, addr: usize, val: i64) {
        let mut buf = self.buffer.lock().unwrap();
        if addr + 8 > buf.len() {
            return;
        }
        buf[addr..addr + 8].copy_from_slice(&val.to_le_bytes());
    }

    pub fn atomic_rmw_add_i64(&self, addr: usize, val: i64) -> i64 {
        let mut buf = self.buffer.lock().unwrap();
        if addr + 8 > buf.len() {
            return 0;
        }
        let old = i64::from_le_bytes(buf[addr..addr + 8].try_into().unwrap());
        buf[addr..addr + 8].copy_from_slice(&old.wrapping_add(val).to_le_bytes());
        old
    }

    pub fn atomic_rmw_sub_i64(&self, addr: usize, val: i64) -> i64 {
        let mut buf = self.buffer.lock().unwrap();
        if addr + 8 > buf.len() {
            return 0;
        }
        let old = i64::from_le_bytes(buf[addr..addr + 8].try_into().unwrap());
        buf[addr..addr + 8].copy_from_slice(&old.wrapping_sub(val).to_le_bytes());
        old
    }

    pub fn atomic_cmpxchg_i64(&self, addr: usize, expected: i64, replacement: i64) -> i64 {
        let mut buf = self.buffer.lock().unwrap();
        if addr + 8 > buf.len() {
            return 0;
        }
        let old = i64::from_le_bytes(buf[addr..addr + 8].try_into().unwrap());
        if old == expected {
            buf[addr..addr + 8].copy_from_slice(&replacement.to_le_bytes());
        }
        old
    }

    // ── Wait / Notify (futex-like) ──────────────────────────────────────
    //
    // WASM spec: memory.atomic.wait32(addr, expected, timeout)
    //   - If memory[addr] != expected, return 1 (not-equal)
    //   - Otherwise, block until notified or timeout
    //   - Return 0 (ok/woken) or 2 (timed-out)
    //
    // WASM spec: memory.atomic.notify(addr, count)
    //   - Wake up to `count` threads waiting on `addr`
    //   - Return number of threads woken

    /// Block until memory[addr] != expected or notified or timeout.
    /// timeout_ns: -1 = infinite, 0 = no wait, >0 = nanoseconds.
    /// Returns: 0 = ok (woken), 1 = not-equal, 2 = timed-out
    pub fn wait32(&self, addr: usize, expected: i32, timeout_ns: i64) -> i32 {
        // Check condition under buffer lock
        {
            let buf = self.buffer.lock().unwrap();
            if addr + 4 > buf.len() {
                return 1;
            }
            let current = i32::from_le_bytes(buf[addr..addr + 4].try_into().unwrap());
            if current != expected {
                return 1; // not-equal
            }
        }

        if timeout_ns == 0 {
            return 2; // timed-out immediately
        }

        let entry = {
            let mut waiters = self.buffer.thread_state.waiters.lock().unwrap();
            waiters
                .entry(addr)
                .or_insert_with(|| Arc::new(WaitEntry::default()))
                .clone()
        };

        let mut state = entry.state.lock().unwrap();
        state.waiters += 1;
        // Scheduler bookkeeping: this thread is parked until wait returns.
        // The guard unparks on EVERY exit path (early returns included).
        self.buffer.thread_state.parked.fetch_add(1, Ordering::SeqCst);
        struct Unpark<'a>(&'a AtomicI32);
        impl Drop for Unpark<'_> {
            fn drop(&mut self) {
                self.0.fetch_sub(1, Ordering::SeqCst);
            }
        }
        let _unpark = Unpark(&self.buffer.thread_state.parked);
        if timeout_ns < 0 {
            while state.notifications == 0 {
                state = entry.condvar.wait(state).unwrap();
            }
            state.notifications -= 1;
            state.waiters -= 1;
            0
        } else {
            let deadline =
                std::time::Instant::now() + std::time::Duration::from_nanos(timeout_ns as u64);
            loop {
                if state.notifications > 0 {
                    state.notifications -= 1;
                    state.waiters -= 1;
                    return 0;
                }
                let now = std::time::Instant::now();
                if now >= deadline {
                    state.waiters -= 1;
                    return 2;
                }
                let timeout = deadline.saturating_duration_since(now);
                let (next_state, result) = entry.condvar.wait_timeout(state, timeout).unwrap();
                state = next_state;
                if result.timed_out() && state.notifications == 0 {
                    state.waiters -= 1;
                    return 2;
                }
            }
        }
    }

    /// Block until memory[addr] != expected or notified or timeout.
    /// timeout_ns: -1 = infinite, 0 = no wait, >0 = nanoseconds.
    /// Returns: 0 = ok (woken), 1 = not-equal, 2 = timed-out
    pub fn wait64(&self, addr: usize, expected: i64, timeout_ns: i64) -> i32 {
        {
            let buf = self.buffer.lock().unwrap();
            if addr + 8 > buf.len() {
                return 1;
            }
            let current = i64::from_le_bytes(buf[addr..addr + 8].try_into().unwrap());
            if current != expected {
                return 1;
            }
        }

        if timeout_ns == 0 {
            return 2;
        }

        let entry = {
            let mut waiters = self.buffer.thread_state.waiters.lock().unwrap();
            waiters
                .entry(addr)
                .or_insert_with(|| Arc::new(WaitEntry::default()))
                .clone()
        };

        let mut state = entry.state.lock().unwrap();
        state.waiters += 1;
        // Scheduler bookkeeping: this thread is parked until wait returns.
        // The guard unparks on EVERY exit path (early returns included).
        self.buffer.thread_state.parked.fetch_add(1, Ordering::SeqCst);
        struct Unpark<'a>(&'a AtomicI32);
        impl Drop for Unpark<'_> {
            fn drop(&mut self) {
                self.0.fetch_sub(1, Ordering::SeqCst);
            }
        }
        let _unpark = Unpark(&self.buffer.thread_state.parked);
        if timeout_ns < 0 {
            while state.notifications == 0 {
                state = entry.condvar.wait(state).unwrap();
            }
            state.notifications -= 1;
            state.waiters -= 1;
            0
        } else {
            let deadline =
                std::time::Instant::now() + std::time::Duration::from_nanos(timeout_ns as u64);
            loop {
                if state.notifications > 0 {
                    state.notifications -= 1;
                    state.waiters -= 1;
                    return 0;
                }
                let now = std::time::Instant::now();
                if now >= deadline {
                    state.waiters -= 1;
                    return 2;
                }
                let timeout = deadline.saturating_duration_since(now);
                let (next_state, result) = entry.condvar.wait_timeout(state, timeout).unwrap();
                state = next_state;
                if result.timed_out() && state.notifications == 0 {
                    state.waiters -= 1;
                    return 2;
                }
            }
        }
    }

    /// Wake up to `count` threads waiting on `addr`.
    /// Returns number of threads actually woken.
    pub fn notify(&self, addr: usize, count: i32) -> i32 {
        if count <= 0 {
            return 0;
        }
        let entry = {
            let waiters = self.buffer.thread_state.waiters.lock().unwrap();
            waiters.get(&addr).cloned()
        };
        let Some(entry) = entry else {
            return 0;
        };
        let woken = {
            let mut state = entry.state.lock().unwrap();
            let available = state.waiters.saturating_sub(state.notifications);
            let woken = available.min(count as usize);
            state.notifications += woken;
            woken
        };
        for _ in 0..woken {
            entry.condvar.notify_one();
        }
        woken as i32
    }
}

impl Default for SharedMemory {
    fn default() -> Self {
        Self::new(0)
    }
}

#[cfg(test)]
mod allocation_bulk_regression_tests {
    use super::{MemoryTrap, SharedMemory};

    #[test]
    fn scalar_accessors_reject_overflow_without_mutation() {
        let memory = SharedMemory::from_vec(vec![0x3c; 16]);
        for addr in [13, 16, usize::MAX - 3, usize::MAX] {
            assert!(matches!(memory.load_i32(addr), Err(MemoryTrap::OutOfBounds {
                addr: bad, size: 4, limit: 16,
            }) if bad == addr));
            assert!(matches!(memory.store_i32(addr, 7), Err(MemoryTrap::OutOfBounds {
                addr: bad, size: 4, limit: 16,
            }) if bad == addr));
            assert!(matches!(memory.load_i64(addr), Err(MemoryTrap::OutOfBounds {
                addr: bad, size: 8, limit: 16,
            }) if bad == addr));
            assert!(matches!(memory.store_i64(addr, 7), Err(MemoryTrap::OutOfBounds {
                addr: bad, size: 8, limit: 16,
            }) if bad == addr));
            assert!(matches!(memory.load_f64(addr), Err(MemoryTrap::OutOfBounds {
                addr: bad, size: 8, limit: 16,
            }) if bad == addr));
            assert!(matches!(memory.store_f64(addr, 7.0), Err(MemoryTrap::OutOfBounds {
                addr: bad, size: 8, limit: 16,
            }) if bad == addr));
        }
        memory.with_buffer(|bytes| assert_eq!(bytes, &[0x3c; 16]));
    }

    #[test]
    fn scalar_accessors_preserve_unaligned_values_and_float_bits() {
        let memory = SharedMemory::new(16);
        memory.store_i32(1, i32::MIN + 17).unwrap();
        assert_eq!(memory.load_i32(1).unwrap(), i32::MIN + 17);
        memory.store_i64(3, i64::MIN + 31).unwrap();
        assert_eq!(memory.load_i64(3).unwrap(), i64::MIN + 31);
        for bits in [0u64, 0x8000_0000_0000_0000, 0x7ff8_0000_dead_beef, 0x7ff0_0000_0000_0000] {
            memory.store_f64(7, f64::from_bits(bits)).unwrap();
            assert_eq!(memory.load_f64(7).unwrap().to_bits(), bits);
        }
    }

    #[test]
    #[ignore = "paired release scalar-accessor measurement"]
    fn scalar_accessor_benchmark() {
        use std::hint::black_box;
        use std::time::Instant;
        fn previous(memory: &SharedMemory, addr: usize) -> Result<i32, MemoryTrap> {
            let buf = memory.buffer.lock().unwrap();
            if addr + 4 > buf.len() {
                return Err(MemoryTrap::OutOfBounds { addr, size: 4, limit: buf.len() });
            }
            Ok(i32::from_le_bytes(buf[addr..addr + 4].try_into().unwrap()))
        }
        fn time(memory: &SharedMemory, old: bool) -> f64 {
            let start = Instant::now();
            for _ in 0..2_000_000 {
                let memory = black_box(memory);
                let addr = black_box(3);
                let value = if old { previous(memory, addr) } else { memory.load_i32(addr) };
                assert_eq!(black_box(value.unwrap()), 0x3c3c3c3c);
            }
            start.elapsed().as_secs_f64()
        }
        let memory = SharedMemory::from_vec(vec![0x3c; 64]);
        let mut ratios = [0.0f64; 5];
        for (pair, ratio) in ratios.iter_mut().enumerate() {
            let (old, new) = if pair % 2 == 0 {
                (time(&memory, true), time(&memory, false))
            } else {
                let new = time(&memory, false);
                (time(&memory, true), new)
            };
            *ratio = old / new;
        }
        ratios.sort_by(f64::total_cmp);
        eprintln!("scalar i32 load: previous/current median {:.3}x; pairs {ratios:?}", ratios[2]);
    }

    #[test]
    fn exclusive_and_shared_growth_match_locked_reference() {
        for shared in [false, true] {
            let mut memory = SharedMemory::new(65536);
            let mut reference = SharedMemory::new(65536);
            memory.set_max_pages(Some(4));
            reference.set_max_pages(Some(4));
            let alias = shared.then(|| memory.clone());
            memory.store_u8(17, 0xa5).unwrap();
            reference.store_u8(17, 0xa5).unwrap();
            for delta in [0, 1, 0, 2, 1, usize::MAX, 0] {
                assert_eq!(memory.grow_exclusive_or_shared(delta), reference.grow(delta));
                reference.with_buffer(|expected| {
                    memory.with_buffer(|actual| assert_eq!(actual, expected));
                });
                if let Some(alias) = &alias {
                    assert_eq!(alias.len(), memory.len());
                }
            }
        }
    }

    #[test]
    #[ignore = "paired release growth-path performance measurement"]
    fn zero_growth_benchmark() {
        use std::hint::black_box;
        use std::time::Instant;
        fn time(memory: &mut SharedMemory, exclusive: bool) -> f64 {
            let start = Instant::now();
            for _ in 0..2_000_000 {
                let memory = black_box(&mut *memory);
                let result = if exclusive {
                    memory.grow_exclusive_or_shared(black_box(0))
                } else {
                    memory.grow(black_box(0))
                };
                black_box(result);
            }
            start.elapsed().as_secs_f64()
        }
        for shared in [false, true] {
            let mut current = SharedMemory::new(65536);
            let mut reference = SharedMemory::new(65536);
            let _current_alias = shared.then(|| current.clone());
            let _reference_alias = shared.then(|| reference.clone());
            let mut ratios = [0.0f64; 21];
            let mut previous_times = [0.0f64; 21];
            let mut current_times = [0.0f64; 21];
            for pair in 0..23 {
                let (old, new) = if pair % 2 == 0 {
                    (time(&mut reference, false), time(&mut current, true))
                } else {
                    let new = time(&mut current, true);
                    (time(&mut reference, false), new)
                };
                if pair >= 2 {
                    let sample = pair - 2;
                    ratios[sample] = old / new;
                    previous_times[sample] = old;
                    current_times[sample] = new;
                }
            }
            assert_eq!(current.len(), reference.len());
            ratios.sort_by(f64::total_cmp);
            previous_times.sort_by(f64::total_cmp);
            current_times.sort_by(f64::total_cmp);
            eprintln!(
                "zero grow shared={shared}: locked/current median {:.3}x; previous_p50_ms={:.3} current_p50_ms={:.3} previous_p95_ms={:.3} current_p95_ms={:.3}; pairs {ratios:?}",
                ratios[10], previous_times[10] * 1000.0, current_times[10] * 1000.0,
                previous_times[19] * 1000.0, current_times[19] * 1000.0,
            );
        }
    }

    #[test]
    #[ignore = "paired loaded-machine nonzero growth measurement"]
    fn nonzero_growth_benchmark() {
        use std::hint::black_box;
        use std::time::Instant;

        fn sample(shared: bool, exclusive: bool) -> f64 {
            // Setup, alias creation, validation and destruction are not timed.
            let mut memories: Vec<_> = (0..128).map(|_| {
                let mut memory = SharedMemory::new(65536);
                memory.set_max_pages(Some(9));
                memory.store_u8(0, 0x91).unwrap();
                memory
            }).collect();
            let aliases = shared.then(|| memories.clone());
            let mut results = [[0usize; 128]; 8];
            let start = Instant::now();
            for round_results in &mut results {
                for (memory, result) in memories.iter_mut().zip(round_results) {
                    let memory = black_box(memory);
                    *result = if exclusive {
                        memory.grow_exclusive_or_shared(black_box(1))
                    } else {
                        memory.grow(black_box(1))
                    };
                }
            }
            let elapsed = start.elapsed().as_secs_f64();
            for (round, results) in results.iter().enumerate() {
                assert!(results.iter().all(|&pages| pages == round + 1));
            }
            for memory in &memories {
                memory.with_buffer(|bytes| {
                    assert_eq!(bytes.len(), 9 * 65536);
                    assert_eq!(bytes[0], 0x91);
                    assert!(bytes[65536..].iter().all(|&byte| byte == 0));
                });
            }
            if let Some(aliases) = aliases {
                assert!(aliases.iter().all(|memory| memory.len() == 9 * 65536));
            }
            elapsed
        }

        for shared in [false, true] {
            let mut ratios = [0.0f64; 21];
            let mut previous_times = [0.0f64; 21];
            let mut current_times = [0.0f64; 21];
            for pair in 0..23 {
                let (old, new) = if pair % 2 == 0 {
                    (sample(shared, false), sample(shared, true))
                } else {
                    let new = sample(shared, true);
                    (sample(shared, false), new)
                };
                if pair >= 2 {
                    let index = pair - 2;
                    ratios[index] = old / new;
                    previous_times[index] = old;
                    current_times[index] = new;
                }
            }
            ratios.sort_by(f64::total_cmp);
            previous_times.sort_by(f64::total_cmp);
            current_times.sort_by(f64::total_cmp);
            eprintln!(
                "eight one-page rounds shared={shared}: locked/current median {:.3}x; previous_p50_ms={:.3} current_p50_ms={:.3} previous_p95_ms={:.3} current_p95_ms={:.3}; pairs {ratios:?}",
                ratios[10], previous_times[10] * 1000.0, current_times[10] * 1000.0,
                previous_times[19] * 1000.0, current_times[19] * 1000.0,
            );
        }
    }

    #[test]
    fn growth_preserves_bytes_zeroes_extension_and_rejects_limits() {
        let mut memory = SharedMemory::new(65536);
        memory.set_max_pages(Some(2));
        memory.store_u8(0, 0x91).unwrap();
        memory.store_u8(65535, 0x42).unwrap();
        let alias = memory.clone();
        assert_eq!(memory.grow(1), 1);
        assert_eq!(alias.len(), 131072);
        memory.with_buffer(|bytes| {
            assert_eq!(bytes[0], 0x91);
            assert_eq!(bytes[65535], 0x42);
            assert!(bytes[65536..].iter().all(|&byte| byte == 0));
        });
        assert_eq!(memory.grow(0), 2);
        assert_eq!(memory.grow(1), usize::MAX);
        assert_eq!(memory.grow(usize::MAX), usize::MAX);
        assert_eq!(alias.len(), 131072);
        assert_eq!(alias.load_u8(0).unwrap(), 0x91);
    }

    #[test]
    fn partial_bulk_helpers_preserve_clipping_and_extreme_offsets() {
        let memory = SharedMemory::from_vec(vec![1, 2, 3, 4]);
        let mut out = [0x99; 8];
        assert_eq!(memory.read_bytes(2, &mut out), 2);
        assert_eq!(out, [3, 4, 0x99, 0x99, 0x99, 0x99, 0x99, 0x99]);
        for addr in [4, 5, usize::MAX] {
            assert_eq!(memory.read_bytes(addr, &mut out), 0);
            memory.write_bytes(addr, &[9, 9]);
        }
        memory.write_bytes(3, &[8, 7, 6]);
        memory.with_buffer(|bytes| assert_eq!(bytes, &[1, 2, 3, 8]));
        assert_eq!(memory.read_bytes(0, &mut []), 0);
        memory.write_bytes(0, &[]);
    }
}

impl SharedMemory {
    /// Decode a checked byte range while its backing is locked. Internal
    /// callers must not reenter memory or invoke guest code from `decode`.
    /// The result must own its data rather than retaining backing pointers.
    pub(crate) fn map_read_range<R>(
        &self,
        addr: usize,
        len: usize,
        decode: impl FnOnce(&[u8]) -> R,
    ) -> Result<R, usize> {
        let buffer = self.buffer.lock().unwrap();
        let limit = buffer.len();
        let end = addr.checked_add(len).filter(|end| *end <= limit).ok_or(limit)?;
        Ok(decode(&buffer[addr..end]))
    }
}

#[cfg(test)]
mod checked_read_range_tests {
    use super::*;

    #[test]
    fn range_is_checked_before_decoder_runs() {
        let memory = SharedMemory::new(16);
        for (addr, len) in [(17, 0), (15, 2), (usize::MAX, 1)] {
            let mut called = false;
            assert_eq!(memory.map_read_range(addr, len, |_| { called = true; }), Err(16));
            assert!(!called);
        }
        assert_eq!(memory.map_read_range(16, 0, |bytes| bytes.len()), Ok(0));
    }

    #[test]
    fn owned_range_survives_alias_growth_and_writes() {
        let memory = SharedMemory::new(65536);
        memory.write_array(3, &[1u8, 2, 3, 255]).unwrap();
        let saved = memory.map_read_range(3, 4, <[u8]>::to_vec).unwrap();
        let alias = memory.clone();
        assert_eq!(alias.grow(1), 1);
        alias.write_array(3, &[9u8; 4]).unwrap();
        assert_eq!(saved, vec![1, 2, 3, 255]);
        assert_eq!(memory.map_read_range(3, 4, <[u8]>::to_vec).unwrap(), vec![9; 4]);
    }
}

impl SharedMemory {
    /// Internal bulk encoders only: no guest calls, memory reentry, or borrowed
    /// pointers may escape the callback. Bounds are checked before mutation.
    pub(crate) fn map_write_range<R>(
        &self,
        addr: usize,
        len: usize,
        encode: impl FnOnce(&mut [u8]) -> R,
    ) -> Result<R, usize> {
        let mut buffer = self.buffer.lock().unwrap();
        let limit = buffer.len();
        let end = addr.checked_add(len).filter(|end| *end <= limit).ok_or(limit)?;
        Ok(encode(&mut buffer[addr..end]))
    }
}

#[cfg(test)]
#[path = "zero_growth_tests.rs"]
mod zero_growth_tests;

#[cfg(test)]
mod exclusive_uniqueness_hint_tests {
    use super::*;

    #[test]
    fn strong_count_hint_never_bypasses_weak_or_shared_ownership() {
        let mut memory = SharedMemory::from_vec(vec![1, 2, 3, 4]);
        let alias = memory.clone();
        assert_eq!(memory.with_buffer_mut_exclusive(|_| panic!("shared closure ran")), None::<()>);
        assert_eq!(alias.with_buffer(|bytes| bytes[0]), 1);
        drop(alias);
        let weak = Arc::downgrade(&memory.buffer);
        assert_eq!(memory.with_buffer_mut_exclusive(|_| panic!("weak alias closure ran")), None::<()>);
        assert!(weak.upgrade().is_some());
        drop(weak);
        assert_eq!(memory.with_buffer_mut_exclusive(|bytes| { bytes[0] = 9; bytes[0] }), Some(9));
        let alias = memory.clone();
        assert_eq!(alias.with_buffer(|bytes| bytes[0]), 9);
        assert_eq!(memory.with_buffer_mut_exclusive(|_| panic!("new shared closure ran")), None::<()>);
    }
}

#[cfg(test)]
mod fixed_access_ownership_guard_tests {
    use std::sync::Arc;

    fn initialized_vm() -> crate::VM {
        let vm = crate::VM::new();
        vm.memory.with_buffer_mut(|buf| {
            buf.clear();
            buf.extend(0..64u8);
            Ok::<(), crate::VMError>(())
        }).unwrap();
        vm
    }

    #[test]
    fn strong_and_weak_aliases_reject_exclusive_access_without_mutation() {
        let mut vm = initialized_vm();
        let expected = std::array::from_fn::<_, 16, _>(|i| i as u8);
        let alias = vm.memory.clone();
        assert_eq!(vm.memory.read_array_exclusive::<16>(0), None);
        assert_eq!(vm.memory.write_array_exclusive(0, &[0xa5; 16]), None);
        assert_eq!(vm.memory.write_u8_checked_exclusive(0, 0xa5), None);
        assert_eq!(vm.memory.with_bytes_exclusive(0, 16, |_| panic!("aliased callback ran")), None::<Result<(), usize>>);
        assert_eq!(alias.read_array::<16>(0), Ok(expected));
        drop(alias);

        let weak = Arc::downgrade(&vm.memory.buffer);
        assert_eq!(vm.memory.read_array_exclusive::<16>(0), None);
        assert_eq!(vm.memory.write_array_exclusive(0, &[0xa5; 16]), None);
        assert_eq!(vm.memory.write_u8_checked_exclusive(0, 0xa5), None);
        assert_eq!(vm.memory.with_bytes_exclusive(0, 16, |_| panic!("weak-aliased callback ran")), None::<Result<(), usize>>);
        let upgraded = weak.upgrade().unwrap();
        assert_eq!(&upgraded.lock().unwrap()[..16], &expected);
        drop(upgraded);
        drop(weak);

        assert_eq!(vm.memory.read_array_exclusive::<16>(0), Some(Ok(expected)));
        assert_eq!(vm.memory.write_array_exclusive(0, &[0xa5; 16]), Some(Ok(())));
        assert_eq!(vm.memory.write_u8_checked_exclusive(63, 0x5a), Some(Ok(())));
        assert_eq!(vm.memory.with_bytes_exclusive(0, 16, |bytes| bytes.len()), Some(Ok(16)));
        assert_eq!(vm.memory.read_array::<16>(0), Ok([0xa5; 16]));
        assert_eq!(vm.memory.read_u8_checked(63), Ok(0x5a));
    }

    fn check_width<const N: usize>() {
        let mut vm = initialized_vm();
        for addr in (0..=65).chain(std::iter::once(usize::MAX)) {
            let before = vm.memory.read_array::<64>(0).unwrap();
            let expected = vm.memory.read_array::<N>(addr);
            assert_eq!(vm.memory.read_array_exclusive::<N>(addr), Some(expected));
            let expected_write = addr.checked_add(N).is_some_and(|end| end <= 64);
            let result = vm.memory.write_array_exclusive(addr, &[0x5a; N]);
            if expected_write {
                assert_eq!(result, Some(Ok(())));
                assert_eq!(vm.memory.read_array::<N>(addr), Ok([0x5a; N]));
            } else {
                assert_eq!(result, Some(Err(64)));
                assert_eq!(vm.memory.read_array::<64>(0), Ok(before));
            }
        }
    }

    #[test]
    fn unique_access_preserves_bounds_empty_ranges_and_unaligned_widths() {
        check_width::<0>();
        check_width::<1>();
        check_width::<2>();
        check_width::<4>();
        check_width::<8>();
        check_width::<16>();
        let mut vm = initialized_vm();
        assert_eq!(vm.memory.write_u8_checked_exclusive(64, 0), Some(Err(64)));
        assert_eq!(vm.memory.write_u8_checked_exclusive(usize::MAX, 0), Some(Err(64)));
        assert_eq!(vm.memory.with_bytes_exclusive(64, 0, |bytes| bytes.len()), Some(Ok(0)));
        assert_eq!(vm.memory.with_bytes_exclusive(65, 0, |_| panic!("out-of-bounds callback ran")), Some(Err(64)));
        assert_eq!(vm.memory.with_bytes_exclusive(usize::MAX, 16, |_| panic!("overflow callback ran")), Some(Err(64)));
    }
}

#[cfg(test)]
mod thread_state_storage_tests {
    use super::*;

    #[test]
    fn cloned_handles_share_counters_and_waiter_entries() {
        let memory = SharedMemory::new(65536);
        let alias = memory.clone();
        assert!(Arc::ptr_eq(&memory.buffer, &alias.buffer));
        assert_eq!(Arc::strong_count(&memory.buffer), 2);
        alias.thread_started();
        assert!(!memory.all_others_parked());
        alias.mark_parked();
        assert!(memory.all_others_parked());
        memory.unmark_parked();
        assert!(!alias.all_others_parked());
        memory.thread_exited();
        assert!(alias.all_others_parked());

        let entry = Arc::new(WaitEntry::default());
        entry.state.lock().unwrap().waiters = 1;
        memory.buffer.thread_state.waiters.lock().unwrap().insert(0, entry.clone());
        assert_eq!(alias.notify(0, 1), 1);
        assert_eq!(entry.state.lock().unwrap().notifications, 1);
        assert_eq!(memory.notify(0, 1), 0);
        alias.store_u8(17, 0xa5).unwrap();
        assert_eq!(memory.load_u8(17).unwrap(), 0xa5);
    }

    #[test]
    fn independently_created_memories_do_not_share_bookkeeping() {
        let first = SharedMemory::new(0);
        let second = SharedMemory::from_vec(Vec::new());
        assert!(!Arc::ptr_eq(&first.buffer, &second.buffer));
        first.thread_started();
        assert!(!first.all_others_parked());
        assert!(second.all_others_parked());
        first.mark_parked();
        assert!(first.all_others_parked());
        assert_eq!(second.buffer.thread_state.parked.load(Ordering::SeqCst), 0);
        first.thread_exited();
        first.unmark_parked();
    }
}

#[cfg(test)]
mod memory_reassignment_tests {
    use super::*;

    #[test]
    #[ignore = "paired shared-memory handle assignment performance measurement"]
    fn same_backing_assignment_benchmark() {
        use std::{hint::black_box, time::Instant};
        fn sample(reuse: bool) -> u128 {
            let source = SharedMemory::new(0);
            let mut destination = source.clone();
            let start = Instant::now();
            for _ in 0..1_000_000 {
                if reuse {
                    black_box(&mut destination).clone_from(black_box(&source));
                } else {
                    *black_box(&mut destination) = black_box(&source).clone();
                }
            }
            let elapsed = start.elapsed().as_nanos();
            assert!(Arc::ptr_eq(&destination.buffer, &source.buffer));
            assert_eq!(Arc::strong_count(&source.buffer), 2);
            elapsed
        }
        let mut ratios = [0.0f64; 11];
        for (pair, ratio) in ratios.iter_mut().enumerate() {
            let (old, new) = if pair % 2 == 0 {
                (sample(false), sample(true))
            } else {
                let new = sample(true);
                (sample(false), new)
            };
            *ratio = old as f64 / new as f64;
        }
        ratios.sort_by(f64::total_cmp);
        eprintln!("same-backing SharedMemory assignment old/current median {:.3}x; pairs {ratios:?}", ratios[5]);
    }

    #[test]
    fn same_backing_assignment_copies_limits_without_detaching_aliases() {
        let mut source = SharedMemory::new(65536);
        source.set_max_pages(Some(2));
        source.store_u8(0, 0x91).unwrap();
        let mut destination = source.clone();
        destination.set_max_pages(Some(1));
        let weak = Arc::downgrade(&source.buffer);
        for _ in 0..32 {
            destination.clone_from(&source);
            assert!(Arc::ptr_eq(&destination.buffer, &source.buffer));
            assert_eq!(Arc::strong_count(&source.buffer), 2);
            assert_eq!(destination.max_pages, Some(2));
            assert!(weak.upgrade().is_some());
        }
        assert_eq!(destination.grow(1), 1);
        assert_eq!(source.len(), 131072);
        assert_eq!(source.load_u8(0).unwrap(), 0x91);
        assert_eq!(destination.grow(1), usize::MAX);
    }

    #[test]
    fn different_backing_assignment_releases_old_family_and_shares_new_one() {
        let source = SharedMemory::new(65536);
        let old = Arc::downgrade(&source.buffer);
        let mut destination = source.clone();
        let mut replacement = SharedMemory::new(65536);
        replacement.set_max_pages(Some(1));
        replacement.store_u8(17, 0xa5).unwrap();
        destination.clone_from(&replacement);
        assert!(Arc::ptr_eq(&destination.buffer, &replacement.buffer));
        assert_eq!(destination.load_u8(17).unwrap(), 0xa5);
        assert_eq!(destination.grow(1), usize::MAX);
        assert_eq!(source.load_u8(17).unwrap(), 0);
        drop(source);
        assert!(old.upgrade().is_none());
        destination.store_u8(17, 0x5a).unwrap();
        assert_eq!(replacement.load_u8(17).unwrap(), 0x5a);
    }
}

#[cfg(test)]
mod scalar_ownership_hint_tests {
    use super::*;
    use crate::VM;
    use std::{hint::black_box, time::Instant};

    #[test]
    fn exclusive_helpers_recover_after_shared_handle_is_dropped() {
        let mut memory = SharedMemory::from_vec(vec![0; 32]);
        let peer = memory.clone();
        assert!(memory.write_u8_checked_exclusive(3, 17).is_none());
        assert!(memory.write_array_exclusive(7, &[1, 2, 3, 4]).is_none());
        assert!(memory.read_u8_checked_exclusive(3).is_none());
        assert!(memory.read_array_exclusive::<4>(7).is_none());
        assert!(memory.with_bytes_exclusive(0, 32, |_| ()).is_none());
        memory.write_array(7, &[1, 2, 3, 4]).unwrap();
        assert_eq!(peer.read_array::<4>(7).unwrap(), [1, 2, 3, 4]);
        drop(peer);
        assert_eq!(memory.read_array_exclusive::<4>(7), Some(Ok([1, 2, 3, 4])));
        assert_eq!(memory.write_u8_checked_exclusive(3, 17), Some(Ok(())));
        assert_eq!(memory.read_u8_checked_exclusive(3), Some(Ok(17)));
        assert_eq!(memory.with_bytes_exclusive(7, 4, |bytes| bytes.to_vec()), Some(Ok(vec![1, 2, 3, 4])));
    }

    #[test]
    fn exclusive_helpers_preserve_checked_bounds_and_do_not_partially_write() {
        let mut memory = SharedMemory::from_vec(vec![7; 32]);
        for address in [32, usize::MAX] {
            assert_eq!(memory.write_u8_checked_exclusive(address, 0), Some(Err(32)));
            assert_eq!(memory.read_u8_checked_exclusive(address), Some(Err(32)));
        }
        for address in [29, usize::MAX] {
            assert_eq!(memory.write_array_exclusive(address, &[0; 4]), Some(Err(32)));
            assert_eq!(memory.read_array_exclusive::<4>(address), Some(Err(32)));
            assert_eq!(memory.with_bytes_exclusive(address, 4, |_| ()), Some(Err(32)));
        }
        assert_eq!(memory.read_array_exclusive::<0>(32), Some(Ok([])));
        assert_eq!(memory.write_array_exclusive(32, &[]), Some(Ok(())));
        assert_eq!(memory.with_buffer(|bytes| bytes.to_vec()), vec![7; 32]);
    }

    // The prior VM memory-zero helpers, retaining the old unconditional
    // Arc::get_mut probe. These are test references, not another runtime path.
    fn legacy(vm: &mut VM, address: usize, value: u8) -> u8 {
        vm.write_memory0_u8_unhinted_reference(address, value).unwrap();
        vm.read_memory0_u8_unhinted_reference(address).unwrap()
    }

    fn locked(vm: &mut VM, address: usize, value: u8) -> u8 {
        vm.write_memory0_u8_locked_reference(address, value).unwrap();
        vm.read_memory0_u8_locked_reference(address).unwrap()
    }

    fn current(vm: &mut VM, address: usize, value: u8) -> u8 {
        vm.write_memory0_u8(address, value).unwrap();
        vm.read_memory0_u8(address).unwrap()
    }

    fn sample(access: fn(&mut VM, usize, u8) -> u8, vm: &mut VM) -> u128 {
        let access = black_box(access);
        let start = Instant::now();
        let mut sum = 0usize;
        for index in 0..100_000 {
            sum += access(black_box(&mut *vm), index & 65535, index as u8) as usize;
        }
        black_box(sum);
        start.elapsed().as_nanos()
    }

    #[test]
    #[ignore = "paired performance experiment; run optimized test executable explicitly"]
    fn benchmark_actual_vm_scalar_memory_helpers() {
        assert!(!cfg!(debug_assertions), "compile with optimizations enabled");
        for shared in [false, true] {
            let mut vm = VM::new();
            vm.memory = SharedMemory::from_vec(vec![0; 65537]);
            let peer = shared.then(|| vm.memory.clone());
            for (candidate, access) in [("hinted", current as fn(&mut VM, usize, u8) -> u8), ("locked", locked)] {
            let mut ratios = Vec::with_capacity(21);
            for round in 0..23 {
                let (before, after) = if round % 2 == 0 {
                    (sample(legacy, &mut vm), sample(access, &mut vm))
                } else {
                    let after = sample(access, &mut vm);
                    (sample(legacy, &mut vm), after)
                };
                assert_eq!(vm.read_memory0_u8(34463).unwrap(), 159);
                assert_eq!(vm.read_memory0_u8(65536).unwrap(), 0);
                if let Some(peer) = &peer {
                    assert_eq!(peer.read_u8_checked(34463).unwrap(), 159);
                    assert_eq!(peer.read_u8_checked(65536).unwrap(), 0);
                }
                if round >= 2 { ratios.push(before as f64 / after as f64); }
            }
            ratios.sort_by(f64::total_cmp);
            println!("VM scalar store/load shared={shared} candidate={candidate} paired_ratio={:.3}", ratios[10]);
            }
        }
    }
}
