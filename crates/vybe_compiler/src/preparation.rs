//! Bounded, persistent workers for preparing independent compilation units.
//! Preparation produces a common AST; linking and execution stay on the caller.

use crate::languages::Language;
use crossbeam_channel::{Receiver, Sender};
use std::path::PathBuf;
use std::sync::atomic::{AtomicBool, Ordering};
use std::sync::{Arc, OnceLock};
use vybe_ast::Module;

pub struct PreparedFile {
    pub path: PathBuf,
    pub source: String,
    pub module: Result<Module, String>,
    pub elapsed: std::time::Duration,
    pub dependencies: Option<Vec<PathBuf>>,
}

impl PreparedFile {
    /// The caller supplies the source it opened at the actual include. A unit
    /// with other dependencies needs separate validation and stays on the
    /// normal preparation path until that validation is available.
    pub fn matches_source(&self, path: &std::path::Path, source: &str) -> bool {
        self.path == path
            && self.source == source
            && self.dependencies.as_ref().is_some_and(Vec::is_empty)
    }
}

/// Request-owned jobs travel with a compiled entry across compiler-service
/// handoffs. Compiled-output caches deliberately do not retain this payload.
pub struct PreparedSources(std::sync::Mutex<std::collections::HashMap<PathBuf, PreparationTicket>>);

impl std::fmt::Debug for PreparedSources {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        f.debug_struct("PreparedSources")
            .field("pending", &self.0.lock().unwrap().len())
            .finish()
    }
}

impl PreparedSources {
    pub(crate) fn new(pending: std::collections::HashMap<PathBuf, PreparationTicket>) -> Self {
        Self(std::sync::Mutex::new(pending))
    }
    pub(crate) fn take(&self) -> std::collections::HashMap<PathBuf, PreparationTicket> {
        std::mem::take(&mut *self.0.lock().unwrap())
    }
}

struct Job {
    path: PathBuf,
    language: Language,
    cancelled: Arc<AtomicBool>,
    result: Sender<Option<PreparedFile>>,
}

pub struct PreparationTicket {
    result: Receiver<Option<PreparedFile>>,
    cancelled: Arc<AtomicBool>,
}

impl PreparationTicket {
    pub fn is_ready(&self) -> bool {
        !self.result.is_empty()
    }

    pub fn wait(self) -> Option<PreparedFile> {
        self.result.recv().ok().flatten()
    }
}

impl Drop for PreparationTicket {
    fn drop(&mut self) {
        self.cancelled.store(true, Ordering::Relaxed);
    }
}

pub struct PreparationPool {
    jobs: Sender<Job>,
}

impl PreparationPool {
    pub fn new(workers: usize, queue_capacity: usize) -> std::io::Result<Self> {
        if workers == 0 || queue_capacity == 0 {
            return Err(std::io::Error::new(
                std::io::ErrorKind::InvalidInput,
                "preparation workers and queue capacity must be positive",
            ));
        }
        let (jobs, receiver) = crossbeam_channel::bounded::<Job>(queue_capacity);
        for index in 0..workers {
            let receiver = receiver.clone();
            std::thread::Builder::new()
                .name(format!("vybe-prepare-{index}"))
                .stack_size(64 * 1024 * 1024)
                .spawn(move || {
                    while let Ok(job) = receiver.recv() {
                        if job.cancelled.load(Ordering::Relaxed) {
                            continue;
                        }
                        let prepared =
                            std::panic::catch_unwind(std::panic::AssertUnwindSafe(|| {
                                let started = std::time::Instant::now();
                                let source = std::fs::read_to_string(&job.path).ok()?;
                                let bundle = crate::dynamic::bundle_from_source(
                                    source,
                                    job.language,
                                    job.path.clone(),
                                );
                                let (module, dependencies) =
                                    crate::bundle::record_source_reads(|| bundle.prepared_module());
                                Some(PreparedFile {
                                    path: job.path,
                                    source: bundle.sources[0].code.clone(),
                                    module,
                                    elapsed: started.elapsed(),
                                    dependencies,
                                })
                            }))
                            .ok()
                            .flatten();
                        let _ = job.result.send(prepared);
                    }
                })?;
        }
        Ok(Self { jobs })
    }

    /// Queue without blocking the executing VM. A full queue simply declines
    /// speculation; the normal synchronous preparation remains available.
    pub fn submit(&self, path: PathBuf, language: Language) -> Option<PreparationTicket> {
        let (sender, result) = crossbeam_channel::bounded(1);
        let cancelled = Arc::new(AtomicBool::new(false));
        self.jobs
            .try_send(Job {
                path,
                language,
                cancelled: cancelled.clone(),
                result: sender,
            })
            .ok()?;
        Some(PreparationTicket { result, cancelled })
    }

    /// Workers share no VM or compilation results. Each ticket owns its data,
    /// so keeping the threads warm does not retain tenant state between runs.
    pub(crate) fn shared() -> Option<&'static Self> {
        static POOL: OnceLock<Option<PreparationPool>> = OnceLock::new();
        POOL.get_or_init(|| Self::new(2, 4).ok()).as_ref()
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use std::sync::{Condvar, Mutex};

    #[test]
    fn independent_units_prepare_concurrently_and_keep_their_results_separate() {
        fn parse(source: &str) -> Result<Module, String> {
            static ARRIVALS: OnceLock<(Mutex<usize>, Condvar)> = OnceLock::new();
            let (lock, ready) = ARRIVALS.get_or_init(|| (Mutex::new(0), Condvar::new()));
            let mut arrived = lock.lock().unwrap();
            *arrived += 1;
            ready.notify_all();
            let (arrived, timeout) = ready
                .wait_timeout_while(arrived, std::time::Duration::from_secs(10), |count| {
                    *count < 2
                })
                .unwrap();
            assert!(
                !timeout.timed_out() && *arrived == 2,
                "both parses must run concurrently"
            );
            Ok(Module {
                name: source.to_string(),
                language: vybe_ast::Lang::Unknown,
                body: vec![],
                imports: vec![],
                directives: Default::default(),
                canon: Default::default(),
            })
        }
        let language = Language {
            name: "preparation-test",
            parse,
            profile_source: || "",
            emit_dispatch: None,
            normalize_class: None,
            register_tree: None,
            expand_source: None,
        };
        let base = std::env::temp_dir().join(format!("vybe-preparation-{}", uuid::Uuid::new_v4()));
        std::fs::create_dir_all(&base).unwrap();
        let first = base.join("first");
        let second = base.join("second");
        std::fs::write(&first, "first unit").unwrap();
        std::fs::write(&second, "second unit").unwrap();
        let pool = PreparationPool::new(2, 4).unwrap();
        let a = pool.submit(first.clone(), language).unwrap();
        let b = pool.submit(second, language).unwrap();
        let a = a.wait().unwrap();
        let b = b.wait().unwrap();
        assert!(a.matches_source(&first, "first unit"));
        assert!(!a.matches_source(&first, "edited unit"));
        assert!(!a.matches_source(&base.join("other"), "first unit"));
        assert_eq!(a.module.unwrap().name, "first unit");
        assert_eq!(b.module.unwrap().name, "second unit");
        assert_eq!(a.source, "first unit");
        assert_eq!(b.source, "second unit");
        std::fs::remove_dir_all(base).unwrap();
    }
}
