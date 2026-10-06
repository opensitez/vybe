//! Actual production JS frontend AST parity, warm allocation counts, and timings.
//! Pass a sample count and --release to time previous/current/native frontends.
//! Includes runtime prelude, UI directives, AST construction and destruction.
//! This measures frontend work; it does not compile AST to WASM.
//! cargo run --features javascript-benchmark --example javascript_ast_parity
use pest_derive::Parser;
use std::{
    alloc::{GlobalAlloc, Layout, System},
    hint::black_box,
    sync::atomic::{AtomicBool, AtomicUsize, Ordering},
    time::Instant,
};
struct CountingAllocator;
static COUNTING: AtomicBool = AtomicBool::new(false);
static ALLOCATIONS: AtomicUsize = AtomicUsize::new(0);
static REALLOCATIONS: AtomicUsize = AtomicUsize::new(0);
#[global_allocator]
static ALLOCATOR: CountingAllocator = CountingAllocator;
unsafe impl GlobalAlloc for CountingAllocator {
    unsafe fn alloc(&self, layout: Layout) -> *mut u8 {
        if COUNTING.load(Ordering::Relaxed) {
            ALLOCATIONS.fetch_add(1, Ordering::Relaxed);
        }
        unsafe { System.alloc(layout) }
    }
    unsafe fn alloc_zeroed(&self, layout: Layout) -> *mut u8 {
        if COUNTING.load(Ordering::Relaxed) {
            ALLOCATIONS.fetch_add(1, Ordering::Relaxed);
        }
        unsafe { System.alloc_zeroed(layout) }
    }
    unsafe fn dealloc(&self, pointer: *mut u8, layout: Layout) {
        unsafe { System.dealloc(pointer, layout) }
    }
    unsafe fn realloc(&self, pointer: *mut u8, layout: Layout, size: usize) -> *mut u8 {
        if COUNTING.load(Ordering::Relaxed) {
            REALLOCATIONS.fetch_add(1, Ordering::Relaxed);
        }
        unsafe { System.realloc(pointer, layout, size) }
    }
}
fn parse_variant(variant: usize, source: &str) -> vybe_ast::Module {
    match variant {
        0 => previous_frontend::parse(source).unwrap(),
        1 => reference_frontend::parse(source).unwrap(),
        _ => native_frontend::parse(source).unwrap(),
    }
}
fn allocation_counts(variant: usize, source: &str) -> (usize, usize) {
    ALLOCATIONS.store(0, Ordering::Relaxed);
    REALLOCATIONS.store(0, Ordering::Relaxed);
    COUNTING.store(true, Ordering::Relaxed);
    let module = parse_variant(variant, source);
    COUNTING.store(false, Ordering::Relaxed);
    black_box(module);
    (
        ALLOCATIONS.load(Ordering::Relaxed),
        REALLOCATIONS.load(Ordering::Relaxed),
    )
}
fn paired_frontends(source: &str, count: usize) -> [(f64, f64); 3] {
    let mut times: [Vec<f64>; 3] = std::array::from_fn(|_| Vec::with_capacity(count));
    for round in 0..count + 10 {
        for index in 0..3 {
            let variant = (index + round) % 3;
            let start = Instant::now();
            black_box(parse_variant(variant, black_box(source)));
            if round >= 10 {
                times[variant].push(start.elapsed().as_secs_f64() * 1e6);
            }
        }
    }
    times.map(|mut times| {
        times.sort_by(f64::total_cmp);
        (times[count / 2], times[count * 95 / 100])
    })
}
#[derive(Parser)]
#[grammar = "../../../languages/js/src/grammar.pest"]
struct ReferenceParser;
// These copies use the production walk and prelude splice, with only Pair's
// backend changed. The compiler's accessor spelling helpers are copied verbatim.
mod class_slots {
    include!(concat!(env!("OUT_DIR"), "/javascript_class_slots.rs"));
}
#[allow(clippy::all)]
mod reference_frontend {
    use super::{ReferenceParser as JsParser, Rule};
    #[allow(dead_code)]
    pub mod walker {
        include!(concat!(env!("OUT_DIR"), "/javascript_reference_walker.rs"));
    }
    include!(concat!(env!("OUT_DIR"), "/javascript_frontend.rs"));
}
#[allow(clippy::all)]
mod previous_frontend {
    use super::{ReferenceParser as JsParser, Rule};
    #[allow(dead_code)]
    pub mod walker {
        include!(concat!(env!("OUT_DIR"), "/javascript_previous_walker.rs"));
    }
    include!(concat!(env!("OUT_DIR"), "/javascript_previous_frontend.rs"));
}
#[allow(dead_code, clippy::upper_case_acronyms)]
mod js {
    include!(concat!(env!("OUT_DIR"), "/javascript_parser.rs"));
}
#[allow(clippy::all)]
mod native_frontend {
    use crate::js::{Parser, Rule};
    struct JsParser;
    impl JsParser {
        fn parse(
            rule: Rule,
            source: &str,
        ) -> Result<grameo::compat::Pairs<'_, Rule>, grameo::ParseError> {
            Parser.parse_pairs(rule, source)
        }
    }
    #[allow(dead_code)]
    pub mod walker {
        include!(concat!(env!("OUT_DIR"), "/javascript_native_walker.rs"));
    }
    include!(concat!(env!("OUT_DIR"), "/javascript_frontend.rs"));
}

fn main() {
    let mut count = None;
    let mut corpus_count = None;
    for argument in std::env::args().skip(1) {
        if let Some(number) = argument.strip_prefix("--corpus=") {
            corpus_count = Some(number.parse::<usize>().expect("corpus file count"));
        } else {
            count = Some(argument.parse::<usize>().expect("sample count"));
        }
    }
    if let Some(count) = count {
        assert!(count > 0);
    }
    for (label, source) in [
        ("tiny", "const x = 1 + 2 * 3;".to_owned()),
        ("declarations100", (0..100).map(|n| format!("const x{n} = {n};\n")).collect()),
        ("expressions100", (0..100).map(|n| format!("const x{n} = f(1 + 2 * 3, (4 - 5) / 6);\n")).collect()),
        ("comments100", (0..100).map(|n| format!("// comment {n}\nconst x{n} = 'é😀';\n")).collect()),
        ("modern", "class Box { constructor(value) { this.value = value; } get() { return this.value; } }\nconst f = (x) => x?.value ?? 0;\nconst xs = [1, 2, 3].map(x => ({ value: x }));\nconst text = `sum: ${f(xs[0])}`;\n".to_owned()),
        ("strict", "/* heading */\n'use strict';\nconst x = 1;".to_owned()),
        ("for_loop", "let x = 0; for (let i = 0; i < 3; i++) { x += i; }".to_owned()),
        ("accessors", "class Box { get value() { return this.x; } set value(v) { this.x = v; } }".to_owned()),
        ("promise_document", "const a = Promise.resolve(document); const b = Promise.reject(document.title); const c = Promise.all([document, Promise.resolve(document.body)]); const d = Promise.any([1, document]);".to_owned()),
        ("promise_without_document", "const a = Promise.resolve(1); const b = Promise.all([1, 2]); const c = Promise.race([3, 4]);".to_owned()),
        ("nested_document", "function f() { if (true) { document.title = 'hi'; } } f();".to_owned()),
        ("document_strings", "const text = 'document'; const object = { document: 1 };".to_owned()),
    ] {
        let reference_module = reference_frontend::parse(&source).unwrap();
        let native_module = native_frontend::parse(&source).unwrap();
        let previous_module = previous_frontend::parse(&source).unwrap();
        assert_eq!(format!("{reference_module:?}"), format!("{native_module:?}"), "{label}: complete common AST including prelude");
        assert_eq!(format!("{reference_module:?}"), format!("{previous_module:?}"), "{label}: prelude allocation common AST parity");
        println!("{label}: common AST including prelude and UI directive identical");
        let counts = std::array::from_fn::<_, 3, _>(|variant| allocation_counts(variant, &source));
        println!("  allocations/reallocations previous {:?}; current {:?}; native {:?}", counts[0], counts[1], counts[2]);
        if let Some(count) = count {
            let [previous, current, native] = paired_frontends(&source, count);
            println!("  frontend median/p95 previous {previous:?} us; current {current:?} us; native {native:?} us; reduction {:.1}%", (1.0 - current.0 / previous.0) * 100.0);
        }
    }
    if let Some(count) = corpus_count {
        check_corpus(count);
    }
}

fn check_corpus(count: usize) {
    fn paths(root: &std::path::Path, output: &mut Vec<std::path::PathBuf>) {
        for entry in std::fs::read_dir(root).unwrap() {
            let path = entry.unwrap().path();
            if path.is_dir() {
                paths(&path, output);
            } else if path.extension().is_some_and(|extension| extension == "js")
                && std::fs::metadata(&path).unwrap().len() <= 8192
            {
                output.push(path);
            }
        }
    }
    assert!(count > 0);
    let root = std::path::Path::new(env!("CARGO_MANIFEST_DIR")).join("../../../tests/js");
    let mut files = Vec::new();
    paths(&root, &mut files);
    files.sort();
    let count = count.min(files.len());
    assert!(count > 0);
    let mut accepted = 0;
    let mut rejected = 0;
    for index in 0..count {
        let path = &files[index * files.len() / count];
        let source = std::fs::read_to_string(path).unwrap();
        // Include invalid-source gates; compare acceptance across both parsers
        // without conflating their intentionally different syntax diagnostics.
        for input in [source.clone(), format!("{source}\0")] {
            let previous = previous_frontend::parse(&input);
            let current = reference_frontend::parse(&input);
            let native = native_frontend::parse(&input);
            match (previous, current, native) {
                (Ok(previous), Ok(current), Ok(native)) => {
                    assert_eq!(
                        format!("{previous:?}"),
                        format!("{current:?}"),
                        "{}: old/current common AST",
                        path.display()
                    );
                    assert_eq!(
                        format!("{current:?}"),
                        format!("{native:?}"),
                        "{}: current/native common AST",
                        path.display()
                    );
                    accepted += 1;
                }
                (Err(previous), Err(current), Err(_)) => {
                    assert_eq!(
                        previous,
                        current,
                        "{}: old/current diagnostics",
                        path.display()
                    );
                    rejected += 1;
                }
                (previous, current, native) => panic!(
                    "{}: acceptance diverged previous={previous:?}; current={current:?}; native={native:?}",
                    path.display()
                ),
            }
        }
    }
    println!(
        "repository JS frontend gate: {count} spread files plus NUL mutations; {accepted} accepted with complete AST parity, {rejected} rejected by all backends"
    );
}
