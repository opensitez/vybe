//! Opt-in compiler-service scheduling evidence, independent of the VM debugger.
//! Times share one monotonic origin; worker durations are not additive request time.
//! With both VYBE_COMPILER_TIMELINE and VYBE_COMPILER_CPU_TIMELINE set, the
//! separate CPU file records wall_ns, thread_id, thread_cpu_ns, phase, event,
//! escaped_path. Union CPU intervals per thread before summing across threads;
//! nested phase totals overlap and must not be added together.
use std::fs::File;
use std::io::Write;
use std::path::{Path, PathBuf};
use std::sync::{Mutex, OnceLock};
use std::time::Instant;

struct Timeline {
    started: Instant,
    file: Mutex<File>,
    cpu_file: Option<Mutex<File>>,
}

fn timeline() -> Option<&'static Timeline> {
    static TRACE: OnceLock<Option<Timeline>> = OnceLock::new();
    TRACE
        .get_or_init(|| {
            let path = std::env::var_os("VYBE_COMPILER_TIMELINE")?;
            if path.is_empty() {
                return None;
            }
            let file = std::fs::OpenOptions::new()
                .create(true)
                .append(true)
                .open(path)
                .ok()?;
            Some(Timeline {
                started: Instant::now(),
                file: Mutex::new(file),
                cpu_file: std::env::var_os("VYBE_COMPILER_CPU_TIMELINE")
                    .filter(|path| !path.is_empty())
                    .and_then(|path| {
                        std::fs::OpenOptions::new()
                            .create(true)
                            .append(true)
                            .open(path)
                            .ok()
                    })
                    .map(Mutex::new),
            })
        })
        .as_ref()
}

fn field(value: &str) -> String {
    value
        .replace('%', "%25")
        .replace('\t', "%09")
        .replace('\n', "%0A")
        .replace('\r', "%0D")
}

impl Timeline {
    fn record(&self, phase: &str, event: &str, path: &Path) {
        let thread = std::thread::current();
        let record = format!(
            "{}\t{}\t{}\t{}\t{}\n",
            self.started.elapsed().as_nanos(),
            field(thread.name().unwrap_or("unnamed")),
            phase,
            event,
            field(&path.to_string_lossy())
        );
        if let Ok(mut file) = self.file.lock() {
            let _ = file.write_all(record.as_bytes());
        }
    }

    fn record_cpu(&self, phase: &str, event: &str, path: &Path, cpu: Option<u64>) {
        let (Some(file), Some(cpu)) = (&self.cpu_file, cpu) else {
            return;
        };
        // Separate file preserves the existing five-column wall trace. CPU
        // coordinates are per-thread, not process CPU (which includes workers).
        let record = format!(
            "{}\t{:?}\t{}\t{}\t{}\t{}\n",
            self.started.elapsed().as_nanos(),
            std::thread::current().id(),
            cpu,
            phase,
            event,
            field(&path.to_string_lossy())
        );
        if let Ok(mut file) = file.lock() {
            let _ = file.write_all(record.as_bytes());
        }
    }
}

fn thread_cpu_ns() -> Option<u64> {
    #[cfg(any(target_os = "macos", target_os = "linux"))]
    {
        // Darwin and Linux use distinct clock IDs. Restrict this ABI to their
        // 64-bit targets; unsupported targets simply omit CPU observations.
        #[cfg(target_pointer_width = "64")]
        {
            unsafe extern "C" {
                fn clock_gettime(clock: i32, value: *mut [i64; 2]) -> i32;
            }
            let mut value = [0_i64; 2];
            let clock = if cfg!(target_os = "macos") { 16 } else { 3 };
            if unsafe { clock_gettime(clock, &mut value) } == 0 {
                return u64::try_from(value[0])
                    .ok()?
                    .checked_mul(1_000_000_000)?
                    .checked_add(u64::try_from(value[1]).ok()?);
            }
        }
    }
    None
}

pub(crate) fn event(phase: &'static str, event: &'static str, path: &Path) {
    if let Some(trace) = timeline() {
        trace.record(phase, event, path);
    }
}

pub(crate) struct Span {
    trace: &'static Timeline,
    phase: &'static str,
    path: PathBuf,
}

pub(crate) fn span(phase: &'static str, path: &Path) -> Option<Span> {
    let trace = timeline()?;
    trace.record(phase, "start", path);
    if trace.cpu_file.is_some() {
        trace.record_cpu(phase, "start", path, thread_cpu_ns());
    }
    Some(Span {
        trace,
        phase,
        path: path.into(),
    })
}

impl Drop for Span {
    fn drop(&mut self) {
        if self.trace.cpu_file.is_some() {
            self.trace
                .record_cpu(self.phase, "end", &self.path, thread_cpu_ns());
        }
        self.trace.record(self.phase, "end", &self.path);
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    #[test]
    fn fields_preserve_record_boundaries_and_literal_escapes() {
        assert_eq!(field("a%09\tb\nc\r"), "a%2509%09b%0Ac%0D");
    }
    #[test]
    fn concurrent_records_remain_complete() {
        let path = std::env::temp_dir().join(format!("vybe-timeline-{}.tsv", std::process::id()));
        let trace = std::sync::Arc::new(Timeline {
            started: Instant::now(),
            file: Mutex::new(File::create(&path).unwrap()),
            cpu_file: None,
        });
        let threads: Vec<_> = (0..2)
            .map(|_| {
                let trace = trace.clone();
                std::thread::spawn(move || {
                    for _ in 0..32 {
                        trace.record("worker", "start", Path::new("a\tb.php"));
                    }
                })
            })
            .collect();
        for thread in threads {
            thread.join().unwrap();
        }
        drop(trace);
        let records = std::fs::read_to_string(&path).unwrap();
        std::fs::remove_file(path).unwrap();
        assert_eq!(records.lines().count(), 64);
        for record in records.lines() {
            let fields: Vec<_> = record.split('\t').collect();
            assert_eq!(fields.len(), 5);
            fields[0].parse::<u128>().unwrap();
            assert_eq!(&fields[2..], &["worker", "start", "a%09b.php"]);
        }
    }

    #[test]
    fn thread_cpu_observations_are_monotonic() {
        if let Some(before) = thread_cpu_ns() {
            let mut value = 1_u64;
            for _ in 0..10_000 {
                value = std::hint::black_box(value.wrapping_mul(3));
            }
            std::hint::black_box(value);
            assert!(thread_cpu_ns().unwrap() >= before);
        }
    }

    #[test]
    fn cpu_records_keep_thread_identity_and_escaped_paths() {
        let path =
            std::env::temp_dir().join(format!("vybe-cpu-timeline-{}.tsv", std::process::id()));
        let wall = path.with_extension("wall");
        let trace = Timeline {
            started: Instant::now(),
            file: Mutex::new(File::create(&wall).unwrap()),
            cpu_file: Some(Mutex::new(File::create(&path).unwrap())),
        };
        trace.record_cpu("emission", "start", Path::new("a\tb.php"), Some(17));
        trace.record_cpu("emission", "end", Path::new("a\tb.php"), Some(29));
        trace.record_cpu("emission", "end", Path::new("ignored"), None);
        drop(trace);
        let records = std::fs::read_to_string(&path).unwrap();
        std::fs::remove_file(path).unwrap();
        std::fs::remove_file(wall).unwrap();
        let rows: Vec<_> = records
            .lines()
            .map(|line| line.split('\t').collect::<Vec<_>>())
            .collect();
        assert_eq!(rows.len(), 2);
        assert_eq!(rows[0].len(), 6);
        assert_eq!(rows[0][1], rows[1][1]);
        assert_eq!(&rows[0][2..], &["17", "emission", "start", "a%09b.php"]);
        assert_eq!(&rows[1][2..], &["29", "emission", "end", "a%09b.php"]);
    }
}
