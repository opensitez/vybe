//! Opt-in linkage attribution, independent of the interactive debugger.

use std::io::Write;
use std::sync::OnceLock;
use std::time::Instant;

pub(crate) struct LinkTiming {
    started: Instant,
    cpu_started: Option<u64>,
    thread: std::thread::ThreadId,
    base: usize,
    chunks: usize,
    bytes: usize,
    linked: bool,
}

impl LinkTiming {
    pub(crate) fn start(base: usize, chunks: &[crate::Chunk]) -> Option<Self> {
        trace_helper_census(base, chunks);
        static ENABLED: OnceLock<bool> = OnceLock::new();
        if !*ENABLED.get_or_init(|| {
            std::env::var_os("VYBE_TRACE_VM_LINK").is_some_and(|v| v == "1")
        }) {
            return None;
        }
        Some(Self {
            started: Instant::now(),
            cpu_started: thread_cpu_ns(),
            thread: std::thread::current().id(),
            base,
            chunks: chunks.len(),
            bytes: chunks.iter().map(|chunk| chunk.code.len()).sum(),
            linked: false,
        })
    }

    pub(crate) fn complete(&mut self) {
        self.linked = true;
    }

    pub(crate) fn stage(&self, phase: &'static str) -> LinkStageTiming {
        LinkStageTiming {
            phase,
            base: self.base,
            thread: self.thread,
            started: Instant::now(),
            cpu_started: thread_cpu_ns(),
        }
    }
}

// This is a diagnostic candidate census, NOT a semantic interning key. Equal
// raw instruction bytes can refer to different constants/imports/directives.
fn helper_census_rows(chunks: &[crate::Chunk])
    -> std::collections::BTreeMap<(&str, u64, usize), usize>
{
    use std::hash::{Hash, Hasher};
    let mut rows = std::collections::BTreeMap::new();
    for chunk in chunks {
        if !chunk.name.starts_with("__") { continue; }
        let mut hash = std::collections::hash_map::DefaultHasher::new();
        chunk.code.hash(&mut hash);
        *rows.entry((chunk.name.as_str(), hash.finish(), chunk.code.len())).or_insert(0) += 1;
    }
    rows
}

fn trace_helper_census(base: usize, chunks: &[crate::Chunk]) {
    use std::sync::atomic::{AtomicU64, Ordering};
    static ENABLED: OnceLock<bool> = OnceLock::new();
    static NEXT_UNIT: AtomicU64 = AtomicU64::new(0);
    if !*ENABLED.get_or_init(|| std::env::var_os("VYBE_TRACE_HELPER_CENSUS")
        .is_some_and(|value| value == "1")) { return; }
    let unit = NEXT_UNIT.fetch_add(1, Ordering::Relaxed);
    let rows = helper_census_rows(chunks);
    let candidates: usize = rows.values().sum();
    let candidate_bytes: usize = rows.iter().map(|((_, _, size), copies)| size * copies).sum();
    let total_bytes: usize = chunks.iter().map(|chunk| chunk.code.len()).sum();
    let mut stderr = std::io::stderr().lock();
    let _ = writeln!(stderr,
        "[vybe-helper-census] unit={} base={} chunks={} ir_bytes={} candidate_chunks={} candidate_ir_bytes={} scope=double-underscore-prefix incoming_attempt=true",
        unit, base, chunks.len(), total_bytes, candidates, candidate_bytes);
    for ((name, hash, size), copies) in rows {
        use std::fmt::Write as _;
        let mut hex = String::with_capacity(name.len() * 2);
        for byte in name.as_bytes() { let _ = write!(hex, "{byte:02x}"); }
        let _ = writeln!(stderr,
            "[vybe-helper-census-body] unit={} name_hex={} raw_hash={:016x} body_ir_bytes={} copies={}",
            unit, hex, hash, size, copies);
    }
}

#[cfg(test)]
mod helper_census_tests {
    use super::*;
    #[test]
    fn groups_raw_candidates_without_claiming_ownership() {
        let mut first = crate::Chunk::new("__stdlib_candidate");
        first.code.extend_from_slice(&[1, 2, 3, 4]);
        let second = first.clone();
        let mut different = first.clone();
        different.code.push(5);
        let ordinary = crate::Chunk::new("ordinary_guest_function");
        // A guest may use this prefix too; the census deliberately makes no
        // ownership claim and must never be used to merge executable bodies.
        let guest = crate::Chunk::new("__guest_candidate");
        let chunks = [first, second, different, ordinary, guest];
        let rows = helper_census_rows(&chunks);
        assert_eq!(rows.len(), 3);
        assert_eq!(rows.values().sum::<usize>(), 4);
        assert!(rows.iter().any(|((name, _, size), copies)|
            *name == "__stdlib_candidate" && *size == 4 && *copies == 2));
    }
}

pub(crate) struct LinkStageTiming {
    phase: &'static str,
    base: usize,
    thread: std::thread::ThreadId,
    started: Instant,
    cpu_started: Option<u64>,
}

impl Drop for LinkStageTiming {
    fn drop(&mut self) {
        let cpu_ended = thread_cpu_ns();
        let wall = self.started.elapsed().as_secs_f64();
        let cpu = self.cpu_started.zip(cpu_ended)
            .and_then(|(start, end)| end.checked_sub(start));
        let start_ns = self.cpu_started.map(|n| n.to_string())
            .unwrap_or_else(|| "unavailable".into());
        let end_ns = cpu_ended.map(|n| n.to_string())
            .unwrap_or_else(|| "unavailable".into());
        let cpu_s = cpu.map(|n| format!("{:.6}", n as f64 / 1e9))
            .unwrap_or_else(|| "unavailable".into());
        let _ = writeln!(std::io::stderr().lock(),
            "[vybe-vm-link-stage] thread={:?} base={} phase={} wall_s={:.6} thread_cpu_s={} thread_cpu_start_ns={} thread_cpu_end_ns={}",
            self.thread, self.base, self.phase, wall, cpu_s, start_ns, end_ns);
    }
}

impl Drop for LinkTiming {
    fn drop(&mut self) {
        let wall = self.started.elapsed().as_secs_f64();
        let cpu_ended = thread_cpu_ns();
        let cpu = self.cpu_started.zip(cpu_ended)
            .and_then(|(start, end)| end.checked_sub(start));
        let start_ns = self.cpu_started.map(|n| n.to_string())
            .unwrap_or_else(|| "unavailable".into());
        let end_ns = cpu_ended.map(|n| n.to_string())
            .unwrap_or_else(|| "unavailable".into());
        let mut stderr = std::io::stderr().lock();
        // A diagnostic must not turn a broken output pipe into a VM panic.
        if let Some(cpu) = cpu {
            let _ = writeln!(stderr,
                "[vybe-vm-link] thread={:?} base={} chunks={} ir_bytes={} linked={} wall_s={:.6} thread_cpu_s={:.6} thread_cpu_start_ns={} thread_cpu_end_ns={}",
                self.thread, self.base, self.chunks, self.bytes, self.linked, wall, cpu as f64 / 1e9, start_ns, end_ns);
        } else {
            let _ = writeln!(stderr,
                "[vybe-vm-link] thread={:?} base={} chunks={} ir_bytes={} linked={} wall_s={:.6} thread_cpu_s=unavailable thread_cpu_start_ns={} thread_cpu_end_ns={}",
                self.thread, self.base, self.chunks, self.bytes, self.linked, wall, start_ns, end_ns);
        }
    }
}

#[cfg(target_os = "macos")]
fn thread_cpu_ns() -> Option<u64> {
    unsafe extern "C" {
        fn clock_gettime_nsec_np(clock: u32) -> u64;
    }
    // CLOCK_THREAD_CPUTIME_ID, verified against the macOS SDK's _time.h.
    Some(unsafe { clock_gettime_nsec_np(16) })
}

#[cfg(all(target_os = "linux", target_pointer_width = "64"))]
fn thread_cpu_ns() -> Option<u64> {
    #[repr(C)]
    struct Timespec { seconds: std::ffi::c_long, nanos: std::ffi::c_long }
    unsafe extern "C" {
        fn clock_gettime(clock: std::ffi::c_int, time: *mut Timespec) -> std::ffi::c_int;
    }
    let mut time = Timespec { seconds: 0, nanos: 0 };
    // Linux CLOCK_THREAD_CPUTIME_ID. Failed observations are not reported as zero.
    if unsafe { clock_gettime(3, &mut time) } != 0 { return None; }
    u64::try_from(time.seconds).ok()?.checked_mul(1_000_000_000)?
        .checked_add(u64::try_from(time.nanos).ok()?)
}

#[cfg(not(any(target_os = "macos", all(target_os = "linux", target_pointer_width = "64"))))]
fn thread_cpu_ns() -> Option<u64> { None }
