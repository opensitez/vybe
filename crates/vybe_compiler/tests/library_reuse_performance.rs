//! Explicit corpus measurement; kept out of ordinary correctness test runs.
#![cfg(target_os = "macos")]
use std::path::Path;
use vybe_compiler::{dynamic::{RuntimeCompilerService, bundle_from_source}, preparation::bind_installed_libraries};

fn cpu_ns() -> u128 {
    unsafe extern "C" { fn clock_gettime(id: i32, time: *mut [i64; 2]) -> i32; }
    let mut time = [0i64; 2];
    // Darwin CLOCK_THREAD_CPUTIME_ID: preparation workers are outside this
    // benchmark; measure the serial compiler and its output teardown.
    assert_eq!(unsafe { clock_gettime(16, &mut time) }, 0);
    time[0] as u128 * 1_000_000_000 + time[1] as u128
}

#[test]
#[ignore = "paired compiler CPU benchmark against local WordPress sources"]
fn paired_wordpress_library_lowering_cpu() {
    vybe_language_php::register();
    let language = vybe_compiler::languages::find_by_name("php").unwrap();
    let mut vm = vybe_runtime::VM::new();
    vybe_compiler::primitives::platforms::register_platforms_all(&mut vm);
    let mut service = RuntimeCompilerService::new(&mut vm);
    service.compile_and_run_source(
        "<?php return (new SplFileInfo('seed.php'))->getFilename();", language, "library-seed.php",
    ).unwrap();
    let root = Path::new(env!("CARGO_MANIFEST_DIR")).join("../../examples/webroot/wordpress/wp-includes");
    for name in ["class-wp-theme-json-resolver.php", "block-template-utils.php"] {
        let path = root.join(name);
        let source = std::fs::read_to_string(&path).unwrap();
        let bundle = bundle_from_source(source, language, &path);
        let original = bundle.prepared_module().unwrap();
        let mut core_only = original.clone();
        let mut all_libraries = original;
        // Before the shared linkage optimization, core exceptions already
        // reused their installed definitions. Keep that advantage in baseline.
        bind_installed_libraries(&mut core_only, |_, identity| identity.starts_with("php:core-exceptions:v1:"));
        bind_installed_libraries(&mut all_libraries, |_, _| true);
        for dependency in vybe_compiler::preparation::required_libraries(&all_libraries) {
            assert!(service.vm().global(&dependency.name).is_some()
                || service.vm().global(&dependency.name.to_ascii_lowercase()).is_some(),
                "benchmark provider {} must be installed", dependency.name);
        }
        let mut baseline = Vec::new();
        let mut reuse = Vec::new();
        let mut sizes = [(0usize, 0usize); 2];
        for round in 0..9 {
            for enabled in if round % 2 == 0 { [false, true] } else { [true, false] } {
                let module = if enabled { &all_libraries } else { &core_only };
                let start = cpu_ns();
                let compiled = bundle.compile_prepared_module(module, &service.vm().modules).unwrap();
                sizes[enabled as usize] = (compiled.chunks.len(), compiled.chunks.iter().map(|chunk| chunk.code.len()).sum());
                std::hint::black_box(&compiled);
                drop(compiled);
                let elapsed = cpu_ns() - start;
                if round != 0 {
                    if enabled { reuse.push(elapsed); } else { baseline.push(elapsed); }
                }
            }
        }
        baseline.sort_unstable(); reuse.sort_unstable();
        assert!(sizes[1].0 < sizes[0].0, "corpus must exercise library code reuse");
        eprintln!("library_lowering file={name} baseline_cpu_ms={:.3} reuse_cpu_ms={:.3} ratio={:.3} baseline_chunks={} reuse_chunks={} baseline_bytes={} reuse_bytes={}",
            baseline[4] as f64 / 1e6, reuse[4] as f64 / 1e6,
            reuse[4] as f64 / baseline[4] as f64, sizes[0].0, sizes[1].0, sizes[0].1, sizes[1].1);
    }
}
