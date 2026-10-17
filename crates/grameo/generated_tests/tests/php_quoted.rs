use grameo::{engine, program::Program, MatchOptions, ParseTree, Span};
use vybe_parser_generated_tests::{php, php_baseline, php_semantic_baseline};

fn captures(tree: &ParseTree<'_, '_>) -> Vec<(String, Span)> {
    let mut out = Vec::new();
    let mut pending: Vec<_> = tree.pairs().collect();
    while let Some(pair) = pending.pop() {
        out.push((pair.rule_name().to_owned(), pair.as_span()));
        pending.extend(pair.into_inner());
    }
    out
}
fn corpus() -> impl Iterator<Item = (&'static str, String)> {
    let root = std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../../examples/webroot/wordpress/wp-includes");
    ["class-wp-hook.php", "class-wp-query.php", "formatting.php"].into_iter()
        .map(move |name| (name, std::fs::read_to_string(root.join(name)).unwrap()))
}

#[test]
fn generated_quote_scanning_preserves_complete_wordpress_capture_trees() {
    assert!(php::Parser.quoted_fragment(php::Parser.rule_id("sq_string").unwrap()).is_some());
    assert!(php_baseline::Parser.quoted_fragment(php_baseline::Parser.rule_id("sq_string").unwrap()).is_none());
    for (name, text) in corpus() {
        let source = text.strip_prefix("<?php").unwrap();
        let actual = engine::parse_program(&php::Parser, "program_pure", source, MatchOptions::default()).unwrap();
        let baseline = engine::parse_program(&php_baseline::Parser, "program_pure", source, MatchOptions::default()).unwrap();
        assert_eq!(captures(&actual), captures(&baseline), "{name}");
    }
}

#[cfg(target_os = "macos")]
#[test]
#[ignore = "paired native PHP grammar CPU benchmark on local WordPress sources"]
fn paired_native_wordpress_quote_scanning_cpu() {
    unsafe extern "C" { fn clock_gettime(id: i32, time: *mut [i64; 2]) -> i32; }
    fn cpu() -> u128 {
        let mut time = [0i64; 2];
        assert_eq!(unsafe { clock_gettime(16, &mut time) }, 0);
        time[0] as u128 * 1_000_000_000 + time[1] as u128
    }
    for (name, text) in corpus() {
        let source = text.strip_prefix("<?php").unwrap();
        let mut times = [Vec::new(), Vec::new()];
        for round in 0..9 {
            for enabled in if round % 2 == 0 { [false, true] } else { [true, false] } {
                let start = cpu();
                let tree = if enabled { engine::parse_program(&php::Parser, "program_pure", source, MatchOptions::default()) }
                    else { engine::parse_program(&php_baseline::Parser, "program_pure", source, MatchOptions::default()) }.unwrap();
                std::hint::black_box(&tree); drop(tree);
                let elapsed = cpu() - start;
                if round != 0 { times[enabled as usize].push(elapsed); }
            }
        }
        for samples in &mut times { samples.sort_unstable(); }
        eprintln!("native_quoted_captures file={name} baseline_cpu_ms={:.3} fast_cpu_ms={:.3} ratio={:.3}",
            times[0][4] as f64 / 1e6, times[1][4] as f64 / 1e6, times[1][4] as f64 / times[0][4] as f64);
    }
}

#[derive(Default)]
struct Hooks(Vec<(usize, Span)>);
impl grameo::builder::Builder for Hooks {
    fn checkpoint(&self) -> usize { self.0.len() }
    fn rollback(&mut self, mark: usize) { self.0.truncate(mark); }
    fn begin(&mut self, _: usize, _: usize) {}
    fn finish(&mut self, rule: usize, span: Span, _: &str) { self.0.push((rule, span)); }
}

#[test]
fn semantic_lexical_scanning_preserves_wordpress_captures_and_hooks() {
    for (name, text) in corpus() {
        let source = text.strip_prefix("<?php").unwrap();
        let mut fast = Hooks::default();
        let mut ordinary = Hooks::default();
        let actual = engine::build_capture_program(&php::Parser, "program_pure", source, MatchOptions::default(), &mut fast).unwrap();
        let baseline = engine::build_capture_program(&php_semantic_baseline::Parser, "program_pure", source, MatchOptions::default(), &mut ordinary).unwrap();
        assert_eq!(captures(&actual), captures(&baseline), "{name}");
        assert_eq!(fast.0, ordinary.0, "{name}");
    }
}

#[cfg(target_os = "macos")]
#[test]
#[ignore = "paired semantic trivia/lookahead CPU benchmark"]
fn paired_native_semantic_lexical_cpu() {
    unsafe extern "C" { fn clock_gettime(id: i32, time: *mut [i64; 2]) -> i32; }
    fn cpu() -> u128 {
        let mut time = [0i64; 2];
        assert_eq!(unsafe { clock_gettime(16, &mut time) }, 0);
        time[0] as u128 * 1_000_000_000 + time[1] as u128
    }
    for (name, text) in corpus() {
        let source = text.strip_prefix("<?php").unwrap();
        let mut times = [Vec::new(), Vec::new()];
        for round in 0..9 {
            for enabled in if round % 2 == 0 { [false, true] } else { [true, false] } {
                let mut hooks = Hooks::default();
                let start = cpu();
                let tree = if enabled { engine::build_capture_program(&php::Parser, "program_pure", source, MatchOptions::default(), &mut hooks) }
                    else { engine::build_capture_program(&php_semantic_baseline::Parser, "program_pure", source, MatchOptions::default(), &mut hooks) }.unwrap();
                std::hint::black_box(&tree); drop(tree); drop(hooks);
                let elapsed = cpu() - start;
                if round != 0 { times[enabled as usize].push(elapsed); }
            }
        }
        for samples in &mut times { samples.sort_unstable(); }
        eprintln!("native_semantic_lexical file={name} baseline_cpu_ms={:.3} fast_cpu_ms={:.3} ratio={:.3}",
            times[0][4] as f64 / 1e6, times[1][4] as f64 / 1e6, times[1][4] as f64 / times[0][4] as f64);
    }
}

#[test]
fn generated_delimited_scanning_preserves_complete_wordpress_capture_trees() {
    use vybe_parser_generated_tests::php_delimited_baseline;
    assert!(php::Parser.delimited_fragment(php::Parser.rule_id("block_comment").unwrap()).is_some());
    assert!(php_delimited_baseline::Parser.delimited_fragment(php_delimited_baseline::Parser.rule_id("block_comment").unwrap()).is_none());
    for (name, text) in corpus() {
        let source = text.strip_prefix("<?php").unwrap();
        let actual = engine::parse_program(&php::Parser, "program_pure", source, MatchOptions::default()).unwrap();
        let baseline = engine::parse_program(&php_delimited_baseline::Parser, "program_pure", source, MatchOptions::default()).unwrap();
        assert_eq!(captures(&actual), captures(&baseline), "{name}");
    }
}

#[cfg(target_os = "macos")]
#[test]
#[ignore = "paired native delimited lexical CPU benchmark"]
fn paired_native_delimited_cpu() {
    use vybe_parser_generated_tests::php_delimited_baseline;
    unsafe extern "C" { fn clock_gettime(id: i32, time: *mut [i64; 2]) -> i32; }
    fn cpu() -> u128 {
        let mut time = [0i64; 2];
        assert_eq!(unsafe { clock_gettime(16, &mut time) }, 0);
        time[0] as u128 * 1_000_000_000 + time[1] as u128
    }
    for (name, text) in corpus() {
        let source = text.strip_prefix("<?php").unwrap();
        let mut times = [Vec::new(), Vec::new()];
        for round in 0..9 {
            for enabled in if round % 2 == 0 { [false, true] } else { [true, false] } {
                let start = cpu();
                let tree = if enabled { engine::parse_program(&php::Parser, "program_pure", source, MatchOptions::default()) }
                    else { engine::parse_program(&php_delimited_baseline::Parser, "program_pure", source, MatchOptions::default()) }.unwrap();
                std::hint::black_box(&tree); drop(tree);
                let elapsed = cpu() - start;
                if round != 0 { times[enabled as usize].push(elapsed); }
            }
        }
        for samples in &mut times { samples.sort_unstable(); }
        eprintln!("native_delimited file={name} baseline_cpu_ms={:.3} fast_cpu_ms={:.3} ratio={:.3}",
            times[0][4] as f64 / 1e6, times[1][4] as f64 / 1e6, times[1][4] as f64 / times[0][4] as f64);
    }
}
