#![cfg(feature = "service-verification")]
//! Same-executable public frontend comparison; no grammar/profile substitution.
use vybe_language_php::{parse, parse_with_peg, with_pratt_admission_cache_for_validation as selected};

fn cpu_ns() -> u128 {
    let mut time = libc::timespec { tv_sec: 0, tv_nsec: 0 };
    assert_eq!(unsafe { libc::clock_gettime(libc::CLOCK_THREAD_CPUTIME_ID, &mut time) }, 0);
    time.tv_sec as u128 * 1_000_000_000 + time.tv_nsec as u128
}
fn sample(source: &str, enabled: bool) -> u128 {
    let start = cpu_ns();
    let module = selected(enabled, || parse(source)).unwrap();
    std::hint::black_box(&module);
    drop(module); // include ordinary AST destruction in both lanes
    cpu_ns() - start
}
fn ast(source: &str, enabled: bool) -> String {
    format!("{:?}", selected(enabled, || parse(source)).unwrap())
}

#[test]
fn public_cache_preserves_complete_ast_errors_and_scopes() {
    for source in [
        "<?php $a = 2; $b = 3; $x = $a + $b * 4; $y = $x << 2; return $y;",
        "<?php $a = 3; function f($x) { global $a; return $x + $a * 2; } return f(4);",
        "<?php $a = 2; $f = function($x) use ($a) { return $x + $a * 3; }; return $f(5);",
        "<?php $a = 1; $b =& $a; $b = $a + 3; return $a;",
        "<?php $a = null; return $a ?? 7 + 2;",
        "<?php return 1 + 2 * 3 . 'x';",
    ] {
        let reference = format!("{:?}", parse_with_peg(source).unwrap());
        assert_eq!(ast(source, false), reference, "uncached {source}");
        assert_eq!(ast(source, true), reference, "cached {source}");
    }
    for source in ["<?php $a + ;", "<?php return $a *;", "<?php $a = 'unterminated"] {
        assert_eq!(selected(false, || parse(source)).unwrap_err(), selected(true, || parse(source)).unwrap_err());
        assert!(parse_with_peg(source).is_err());
    }

}

#[test]
fn wordpress_public_frontend_paired_debug_cpu() {
    assert!(cfg!(debug_assertions), "use the debug build");
    let root = std::path::Path::new(env!("CARGO_MANIFEST_DIR")).join("../..");
    for name in ["functions.php", "class-wp-query.php", "formatting.php", "script-loader.php"] {
        let path = root.join("examples/webroot/wordpress/wp-includes").join(name);
        let source = std::fs::read_to_string(path).unwrap();
        let before = ast(&source, false);
        let after = ast(&source, true);
        assert_eq!(before, after, "same optimized public parser, full Module {name}");
        assert_eq!(after, format!("{:?}", parse_with_peg(&source).unwrap()), "PEG Module {name}");
        drop((before, after));
        sample(&source, false); sample(&source, true);
        let mut baseline = Vec::new(); let mut cached = Vec::new();
        for pair in 0..6 {
            for enabled in if pair % 2 == 0 { [false, true] } else { [true, false] } {
                let cpu = sample(&source, enabled);
                if enabled { cached.push(cpu); } else { baseline.push(cpu); }
            }
        }
        baseline.sort_unstable(); cached.sort_unstable();
        println!("pratt_admission_cpu file={name} bytes={} baseline_ns={baseline:?} cached_ns={cached:?} median_baseline_ns={} median_cached_ns={}", source.len(), (baseline[2]+baseline[3])/2, (cached[2]+cached[3])/2);
    }
}
