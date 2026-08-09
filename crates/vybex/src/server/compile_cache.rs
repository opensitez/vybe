//! Compiled-output cache for `--serve`.
//!
//! A server answers requests for the SAME file over and over, and compiling it
//! is not a rounding error: `--dump` (compile, no run) costs the same CPU as a
//! full run, so essentially the whole per-request cost is compilation, and it
//! grows super-linearly with source size. Caching the compiled output is worth
//! more here than any amount of tuning inside the compiler.
//!
//! Measured side by side, same machine and same window, `--pool 4` with and
//! without this cache:
//!
//! | PHP page | no cache | cached |
//! |----------|----------|--------|
//! | 0.1 KB   | 22 ms    | 5 ms   |
//! | 0.5 KB   | 35 ms    | 11 ms  |
//! | 5 KB     | 320 ms   | 57 ms  |
//!
//! Absolute numbers move a lot with machine load — the same 5 KB page measured
//! ~1.5s under heavy load — so only same-window pairs like the table above mean
//! anything. The ratio is the durable part.
//!
//! What is cached is the output of `RuntimeCompilerService::compile_bundle` —
//! chunks plus host-import metadata — and a request on a hit does only
//! `run_compiled`.
//!
//! ## Two things make this legal, and neither is obvious
//!
//! 1. **Chunks are position-independent.** `run_compiled_impl` takes
//!    `base_chunk_index = vm.chunks.len()` and relocates chunk indices, import
//!    tables and type tables against it. Without that, cached chunks compiled
//!    against one VM state could not be installed into another.
//! 2. **Every pool VM resets to the same baseline.** `compile_bundle` compiles
//!    against `vm.modules`, so a cached compilation is only valid while that
//!    map is what it was at boot. It is — because `reset_to` restores it before
//!    every request. The cache's correctness therefore RESTS on the pool's
//!    reset guarantee; weaken one and the other stops holding.
//!
//! ## Invalidation, and why the dependency set has to come from the compiler
//!
//! Keying on the entry file's mtime is not enough and the failure is silent.
//! `Bundle::prepared_module` resolves source imports at COMPILE time — PHP
//! `require_once`, and per its own doc any include/import a front-end can
//! resolve statically — so the compile reads files that never appear in
//! `bundle.sources`. Demonstrated: a `page.php` that requires `lib.php` changes
//! its output when only `lib.php` is edited, while `page.php`'s own mtime and
//! size are byte-for-byte unchanged. An entry-mtime cache serves the old page
//! forever.
//!
//! So the dependency set comes from the compiler itself:
//! [`compile_with_dependencies`] runs the compile inside
//! `vybe_compiler::bundle::record_source_reads`, which records every file the
//! preparation path opens. [`CompileCache::store`] refuses to cache at all when
//! that set is not known to be complete — better to cache nothing than to serve
//! stale code, and a cache that is wrong on exactly the files a real
//! application is built from is not a cache worth having.

use std::path::{Path, PathBuf};
use std::sync::Arc;
use std::time::SystemTime;

use vybe_compiler::dynamic::DynamicCompilation;
use vybe_compiler::primitives::HostImportMetadata;
use vybe_runtime::Chunk;

/// One file the compilation read, and the stamp that says it hasn't changed.
///
/// `(mtime, len)` rather than a content hash: a hit must be cheap or it eats
/// the win, and an editor writing a same-length file within one mtime tick is
/// the only miss — which `--no-cache` covers.
#[derive(Debug, Clone)]
struct Dep {
    path: PathBuf,
    /// `false` when the compile looked for this file and did not find it. That
    /// is a real dependency: PHP's optional-include path compiles differently
    /// depending on the absence, so a file APPEARING has to invalidate exactly
    /// as an edit does — and "absent then, absent now" has to stay a hit, or an
    /// app with one optional include would never see a cache hit at all.
    present: bool,
    mtime: Option<SystemTime>,
    len: u64,
}

impl Dep {
    fn stamp(path: PathBuf) -> Self {
        match std::fs::metadata(&path) {
            Ok(m) => Dep {
                present: true,
                mtime: m.modified().ok(),
                len: m.len(),
                path,
            },
            Err(_) => Dep {
                present: false,
                mtime: None,
                len: 0,
                path,
            },
        }
    }

    fn unchanged(&self) -> bool {
        match std::fs::metadata(&self.path) {
            Ok(m) => self.present && m.len() == self.len && m.modified().ok() == self.mtime,
            // Gone. A hit only if it was already gone when we compiled; a file
            // deleted since then changes the answer.
            Err(_) => !self.present,
        }
    }
}

/// A compile FAILURE is cached too, under the same dependency set.
///
/// Not an optimisation — a pool-slot protection. A page that fails to compile
/// costs exactly as much to fail as to succeed (a 21 KB PHP file measured
/// ~5.8s), and recompiling it on every request would burn a warm VM slot for
/// that long each time, for as long as the file stays broken; enough of them
/// wedge the whole pool. Keying the error on the same files means the developer
/// still gets a fresh compile the moment they touch any of them — which is the
/// only time the error can have changed.
enum Outcome {
    Compiled {
        chunks: Vec<Chunk>,
        host_imports: HostImportMetadata,
        entry_path: Option<PathBuf>,
    },
    Failed(String),
}

struct Entry {
    deps: Vec<Dep>,
    outcome: Outcome,
}

#[derive(Default)]
pub struct CompileCache {
    entries: dashmap::DashMap<PathBuf, Arc<Entry>>,
    /// Runtime `include`/`require` compilations, which the entry-level map
    /// above cannot hold: they happen during `run_compiled`, long after
    /// `compile_bundle` returned, and the same file compiles differently
    /// against a different module map or a different entry — hence the
    /// fingerprint in the key rather than a second validity rule.
    includes: dashmap::DashMap<(PathBuf, u64), Arc<Entry>>,
    hits: std::sync::atomic::AtomicU64,
    misses: std::sync::atomic::AtomicU64,
}

impl CompileCache {
    pub fn new() -> Self {
        Self::default()
    }

    /// The cached result for `key` — a ready-to-run compilation or the compile
    /// error it produced — if every file it was built from is unchanged.
    pub fn get(&self, key: &Path) -> Option<Result<DynamicCompilation, String>> {
        use std::sync::atomic::Ordering;
        let entry = self.entries.get(key)?;
        if !entry.deps.iter().all(Dep::unchanged) {
            drop(entry);
            // Drop it rather than leave it to be re-checked: a file being
            // edited will keep changing, and a stale entry that fails
            // validation on every request is pure cost.
            self.entries.remove(key);
            self.misses.fetch_add(1, Ordering::Relaxed);
            return None;
        }
        self.hits.fetch_add(1, Ordering::Relaxed);
        Some(match &entry.outcome {
            Outcome::Compiled {
                chunks,
                host_imports,
                entry_path,
            } => Ok(DynamicCompilation {
                chunks: chunks.clone(),
                host_imports: host_imports.clone(),
                entry_path: entry_path.clone(),
            }),
            Outcome::Failed(message) => Err(message.clone()),
        })
    }

    /// Cache the outcome of compiling `key`, success or failure.
    ///
    /// `deps` must be EVERY file the compilation read, entry file included.
    /// `None` means the compiler could not say — see the module header — and
    /// nothing is cached at all.
    pub fn store(
        &self,
        key: &Path,
        deps: Option<Vec<PathBuf>>,
        result: Result<&DynamicCompilation, &str>,
    ) {
        use std::sync::atomic::Ordering;
        self.misses.fetch_add(1, Ordering::Relaxed);
        let Some(deps) = deps else {
            return;
        };
        let outcome = match result {
            Ok(compiled) => Outcome::Compiled {
                chunks: compiled.chunks.clone(),
                host_imports: compiled.host_imports.clone(),
                entry_path: compiled.entry_path.clone(),
            },
            Err(message) => Outcome::Failed(message.to_string()),
        };
        self.entries.insert(
            key.to_path_buf(),
            Arc::new(Entry {
                deps: deps.into_iter().map(Dep::stamp).collect(),
                outcome,
            }),
        );
    }

    /// `(hits, misses, entries)` — for the shutdown line and for tests.
    pub fn stats(&self) -> (u64, u64, usize) {
        use std::sync::atomic::Ordering;
        (
            self.hits.load(Ordering::Relaxed),
            self.misses.load(Ordering::Relaxed),
            self.entries.len() + self.includes.len(),
        )
    }
}

/// The compiled output of a cached entry, or the error it produced.
fn replay(outcome: &Outcome) -> Result<DynamicCompilation, String> {
    match outcome {
        Outcome::Compiled {
            chunks,
            host_imports,
            entry_path,
        } => Ok(DynamicCompilation {
            // Cloned, never handed out by reference: installing a compilation
            // RELOCATES its chunk indices against the VM it is going into, so a
            // shared copy would be rewritten by its first use.
            chunks: chunks.clone(),
            host_imports: host_imports.clone(),
            entry_path: entry_path.clone(),
        }),
        Outcome::Failed(message) => Err(message.clone()),
    }
}

fn outcome_of(result: Result<&DynamicCompilation, &str>) -> Outcome {
    match result {
        Ok(compiled) => Outcome::Compiled {
            chunks: compiled.chunks.clone(),
            host_imports: compiled.host_imports.clone(),
            entry_path: compiled.entry_path.clone(),
        },
        Err(message) => Outcome::Failed(message.to_string()),
    }
}

/// Runtime includes reuse the entry cache's storage, validation and counters —
/// only the key differs. Keeping them in one place is what makes "outside the
/// VM, cleared with the server, never shared between tenants" true of both
/// without a second thing to reason about.
impl vybe_compiler::dynamic::IncludeCompileCache for CompileCache {
    fn get(&self, path: &Path, fingerprint: u64) -> Option<Result<DynamicCompilation, String>> {
        use std::sync::atomic::Ordering;
        let key = (path.to_path_buf(), fingerprint);
        let entry = self.includes.get(&key)?;
        if !entry.deps.iter().all(Dep::unchanged) {
            drop(entry);
            self.includes.remove(&key);
            self.misses.fetch_add(1, Ordering::Relaxed);
            return None;
        }
        self.hits.fetch_add(1, Ordering::Relaxed);
        Some(replay(&entry.outcome))
    }

    fn store(
        &self,
        path: &Path,
        fingerprint: u64,
        deps: Option<Vec<PathBuf>>,
        result: Result<&DynamicCompilation, &str>,
    ) {
        use std::sync::atomic::Ordering;
        self.misses.fetch_add(1, Ordering::Relaxed);
        // `None` means the compiler could not attribute the reads — cache
        // nothing rather than serve an include whose edits go unnoticed.
        let Some(deps) = deps else {
            return;
        };
        self.includes.insert(
            (path.to_path_buf(), fingerprint),
            Arc::new(Entry {
                deps: deps.into_iter().map(Dep::stamp).collect(),
                outcome: outcome_of(result),
            }),
        );
    }
}

/// Compile `bundle`, and report every file the compilation read.
///
/// The two halves have to happen together: the dependency set is collected
/// *during* the compile by `vybe_compiler::bundle::record_source_reads`, which
/// is the only thing that can answer completely — the reads happen inside
/// `prepared_module`, at the two `read_to_string` sites that module marks.
///
/// Deliberately NOT approximated on this side. Re-deriving "which files does
/// this script include" in the server would be a second implementation of
/// include resolution living beside the real one, and the two would disagree —
/// the failure mode the request-superglobal work already had to undo once.
///
/// `None` means the reads could not be attributed to this compile (a nested
/// compilation), and the caller must then not cache.
pub fn compile_with_dependencies(
    compiler: &mut vybe_compiler::dynamic::RuntimeCompilerService<'_>,
    bundle: &vybe_compiler::bundle::Bundle,
) -> (Result<DynamicCompilation, String>, Option<Vec<PathBuf>>) {
    let (result, reads) =
        vybe_compiler::bundle::record_source_reads(|| compiler.compile_bundle(bundle));
    // The entry file (and any other source the LOADER opened) is read before
    // the compile starts, so it is never in `reads` — union it in, or editing
    // the page itself would not invalidate anything.
    let deps = reads.map(|reads| {
        let mut all: Vec<PathBuf> = bundle.sources.iter().map(|s| s.path.clone()).collect();
        for path in reads {
            if !all.contains(&path) {
                all.push(path);
            }
        }
        all
    });
    (result, deps)
}

#[cfg(test)]
mod tests {
    use super::*;

    fn empty_compilation() -> DynamicCompilation {
        DynamicCompilation {
            chunks: vec![Chunk::new("t".to_string())],
            host_imports: HostImportMetadata::default(),
            entry_path: None,
        }
    }

    fn temp_file(name: &str, body: &str) -> PathBuf {
        let path = std::env::temp_dir().join(format!("vybe_cc_{name}"));
        std::fs::write(&path, body).unwrap();
        path
    }

    #[test]
    fn a_compilation_with_an_unknown_dependency_set_is_never_cached() {
        let cache = CompileCache::new();
        let entry = temp_file("unknown.php", "<?php echo 1;");
        cache.store(&entry, None, Ok(&empty_compilation()));
        assert!(
            cache.get(&entry).is_none(),
            "storing with deps=None must cache nothing — better a miss than stale code"
        );
        assert_eq!(cache.stats().2, 0);
    }

    #[test]
    fn an_unchanged_dependency_set_hits() {
        let cache = CompileCache::new();
        let entry = temp_file("hit.php", "<?php echo 1;");
        cache.store(&entry, Some(vec![entry.clone()]), Ok(&empty_compilation()));
        assert!(matches!(cache.get(&entry), Some(Ok(_))));
        assert_eq!(cache.stats().0, 1, "one hit");
    }

    /// The whole reason the dependency set has to be complete: the entry file
    /// is untouched and only an INCLUDED file changed.
    #[test]
    fn editing_a_dependency_that_is_not_the_entry_file_invalidates() {
        let cache = CompileCache::new();
        let entry = temp_file("page_dep.php", "<?php require_once 'libdep.php';");
        let lib = temp_file("libdep.php", "<?php function b() { return 1; }");
        cache.store(
            &entry,
            Some(vec![entry.clone(), lib.clone()]),
            Ok(&empty_compilation()),
        );
        assert!(cache.get(&entry).is_some(), "fresh entry hits");

        // Same trick a developer's editor performs: rewrite the include only.
        std::thread::sleep(std::time::Duration::from_millis(1100));
        std::fs::write(&lib, "<?php function b() { return 2; }").unwrap();

        assert!(
            cache.get(&entry).is_none(),
            "a changed include must invalidate even though the entry file did not move"
        );
        assert_eq!(
            cache.stats().2,
            0,
            "the dead entry is dropped, not re-checked"
        );
    }

    /// A broken page must not cost a full compile per request — see
    /// [`Outcome::Failed`]. The error comes back from cache, and touching the
    /// file is what earns a fresh one.
    #[test]
    fn a_compile_error_is_cached_and_re_reported_without_recompiling() {
        let cache = CompileCache::new();
        let entry = temp_file("broken.php", "<?php this is not php");
        cache.store(
            &entry,
            Some(vec![entry.clone()]),
            Err("compile error: unexpected token"),
        );
        match cache.get(&entry) {
            Some(Err(message)) => assert!(message.contains("unexpected token")),
            other => panic!("expected the cached compile error, got {:?}", other.is_some()),
        }

        std::thread::sleep(std::time::Duration::from_millis(1100));
        std::fs::write(&entry, "<?php echo 1;").unwrap();
        assert!(
            cache.get(&entry).is_none(),
            "fixing the file must earn a fresh compile"
        );
    }

    #[test]
    fn a_deleted_dependency_invalidates() {
        let cache = CompileCache::new();
        let entry = temp_file("gone.php", "<?php echo 1;");
        cache.store(&entry, Some(vec![entry.clone()]), Ok(&empty_compilation()));
        std::fs::remove_file(&entry).unwrap();
        assert!(cache.get(&entry).is_none());
    }
}
