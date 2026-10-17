//! Explicit local corpus timing; never run as part of ordinary correctness tests.
use std::path::Path;

fn direct_parse(source: &str) -> Result<vybe_ast::Module, String> {
    if std::env::var_os("GRAMEO_PHP_PRATT_ONLY").is_some() {
        vybe_language_php::walker::parse_direct_pratt(source)
    } else {
        vybe_language_php::walker::parse_direct(source)
    }
}

fn cpu_ns() -> u128 {
    let mut time = libc::timespec { tv_sec: 0, tv_nsec: 0 };
    assert_eq!(unsafe { libc::clock_gettime(libc::CLOCK_PROCESS_CPUTIME_ID, &mut time) }, 0);
    time.tv_sec as u128 * 1_000_000_000 + time.tv_nsec as u128
}

#[test]
#[ignore = "paired process CPU benchmark against local WordPress sources"]
fn paired_php_parse_and_ast_cpu() {
    let root = Path::new(env!("CARGO_MANIFEST_DIR")).join("../../examples/webroot/wordpress/wp-includes");
    for name in ["class-wp-hook.php", "class-wp-query.php", "formatting.php"] {
        let source = std::fs::read_to_string(root.join(name)).unwrap();
        // Warm bootstrap caches before alternating orders on the busy machine.
        std::hint::black_box(vybe_language_php::parse_with_peg(&source).unwrap());
        std::hint::black_box(direct_parse(&source).unwrap());
        let mut baseline = Vec::new();
        let mut direct = Vec::new();
        for round in 0..9 {
            for enabled in if round % 2 == 0 { [false, true] } else { [true, false] } {
                let start = cpu_ns();
                let module = if enabled { direct_parse(&source) }
                    else { vybe_language_php::parse_with_peg(&source) }.unwrap();
                std::hint::black_box(&module);
                drop(module);
                let elapsed = cpu_ns() - start;
                if enabled { direct.push(elapsed); } else { baseline.push(elapsed); }
            }
        }
        baseline.sort_unstable();
        direct.sort_unstable();
        eprintln!("php_parse_and_ast file={name} bytes={} baseline_cpu_ms={:.3} direct_cpu_ms={:.3} ratio={:.3}",
            source.len(), baseline[4] as f64 / 1e6, direct[4] as f64 / 1e6,
            direct[4] as f64 / baseline[4] as f64);
    }
}

#[test]
#[ignore = "complete local WordPress corpus AST compatibility gate"]
fn wordpress_direct_ast_corpus_matches_default_frontend() {
    fn visit(path: &Path, files: &mut Vec<std::path::PathBuf>) {
        let mut entries = std::fs::read_dir(path).unwrap()
            .map(|entry| entry.unwrap()).collect::<Vec<_>>();
        entries.sort_by_key(|entry| entry.file_name());
        for entry in entries {
            let kind = entry.file_type().unwrap();
            if kind.is_dir() { visit(&entry.path(), files); }
            else if kind.is_file() && entry.path().extension().is_some_and(|ext| ext == "php") {
                files.push(entry.path());
            }
        }
    }
    std::thread::Builder::new().stack_size(32 * 1024 * 1024).spawn(|| {
        let root = Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("../../examples/webroot/wordpress");
        let mut files = Vec::new();
        visit(&root, &mut files);
        assert!(!files.is_empty(), "WordPress corpus is missing");
        // Optional exact-file selection makes a failing phase reproducible
        // with this same executable. The default gate still runs every file.
        let filter = std::env::var_os("GRAMEO_WORDPRESS_CORPUS_FILE");
        if let Some(filter) = &filter {
            files.retain(|path| path.strip_prefix(&root).ok() == Some(Path::new(filter)));
            assert!(!files.is_empty(), "selected WordPress corpus file is missing");
        }
        let (mut parsed, mut rejected) = (0usize, 0usize);
        for path in &files {
            let source = std::fs::read_to_string(path).unwrap();
            eprintln!("corpus_phase default_parse {}", path.display());
            let baseline = vybe_language_php::parse_with_peg(&source);
            eprintln!("corpus_phase direct_parse {}", path.display());
            let direct = direct_parse(&source);
            eprintln!("corpus_phase compare {}", path.display());
            match (baseline, direct) {
                (Ok(baseline), Ok(direct)) => {
                    // Debug covers the full common AST, including all spans.
                    // Report a compact mismatch, never dump multi-MB trees.
                    let baseline = format!("{baseline:?}");
                    let direct = format!("{direct:?}");
                    if baseline != direct {
                        let offset = baseline.chars().zip(direct.chars())
                            .position(|(a, b)| a != b).unwrap_or_else(||
                                baseline.chars().count().min(direct.chars().count()));
                        panic!("AST/span mismatch in {} at character {offset}", path.display());
                    }
                    parsed += 1;
                    eprintln!("corpus_phase complete {}", path.display());
                }
                (Err(baseline), Err(direct)) => {
                    assert!(baseline == direct, "parse diagnostics differ in {}", path.display());
                    rejected += 1;
                }
                _ => panic!("parse pass/fail differs in {}", path.display()),
            }
            eprintln!("corpus_phase dropped {}", path.display());
        }
        eprintln!("wordpress_direct_ast mode={} files={} parsed={parsed} rejected={rejected} mismatches=0",
            if filter.is_some() { "single-file" } else { "full-corpus" }, files.len());
    }).unwrap().join().unwrap();
}
