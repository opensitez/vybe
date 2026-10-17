//! Debug-only paired emission gate. Parse/AST comparison and output validation
//! are outside timing. Timed samples include Compiler construction, emission,
//! finalization, linked-code preparation and destruction on the current thread.
//! This is not a complete WordPress request or detached CaaS-worker benchmark.
use std::{sync::Arc, time::Instant};
use vybe_compiler::Compiler;
use vybe_runtime::{Chunk, PreparedLinkedCode, VM, import_table::{ImportInterningGuard, ImportInterningMode}};

fn cpu_ns() -> u128 {
    unsafe extern "C" { fn clock_gettime(id: i32, time: *mut [i64; 2]) -> i32; }
    let mut time = [0; 2];
    #[cfg(target_os = "macos")] let clock = 16;
    #[cfg(not(target_os = "macos"))] let clock = 3;
    assert_eq!(unsafe { clock_gettime(clock, &mut time) }, 0);
    time[0] as u128 * 1_000_000_000 + time[1] as u128
}

#[path = "emission_import_interning/canonical.rs"]
mod metadata;
use metadata::canonical;

fn compile(module: &vybe_ast::Module, profile: &Arc<vybe_runtime::profile::LanguageProfile>, options: vybe_runtime::registry::CompilerOptimizations, mode: ImportInterningMode) -> Vec<Chunk> {
    let _guard = ImportInterningGuard::enter(mode);
    Compiler::with_shared_profile(profile.clone()).with_optimizations(options)
        .with_helper_templates(vybe_language_php::parse,
            vybe_language_php::emitter::serialization_adapter::HELPER_DEFINITIONS,
            Arc::new(vybe_compiler::primitives::helper_templates::Pool::default()))
        .with_argument_emitters(vybe_language_php::parse,
            vybe_language_php::emitter::argument_emission::ARGUMENT_DEFINITIONS)
        .with_run_module_end_hooks(false).compile(module).unwrap()
}

fn same(a: &[Chunk], b: &[Chunk], unit: &str) {
    assert_eq!(a.len(), b.len(), "chunk count {unit}");
    for (i, (a, b)) in a.iter().zip(b).enumerate() {
        assert!(a.code == b.code, "code {unit} chunk {i}");
        assert!(a.lines == b.lines, "source lines {unit} chunk {i}");
        assert_eq!(a.imports, b.imports, "imports {unit} chunk {i}");
        assert_eq!(a.global_imports, b.global_imports, "global imports {unit} chunk {i}");
        assert!(canonical(&format!("{a:?}")) == canonical(&format!("{b:?}")), "complete metadata {unit} chunk {i}");
    }
    let write = |chunks: &[Chunk]| vybe_platform_wasm::write_wasm_with_options(chunks,
        vybe_platform_wasm::WasmWriteOptions { include_vybe_metadata: false, include_name_section: true });
    let wa = write(a); let wb = write(b);
    assert!(wa == wb, "complete standard Wasm bytes {unit}");
    let validation = wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all()).validate_all(&wa).map(|_| ());
    match validation {
        Ok(()) => println!("parity unit={unit} chunks={} bytes={} wasm_bytes={} validation=pass", a.len(), a.iter().map(|c| c.code.len()).sum::<usize>(), wa.len()),
        Err(error) => {
            // This is a shared baseline limit, not a successful validation.
            // Continue only for wasmparser's explicit implementation size cap.
            let message = error.to_string();
            assert!(message.contains("function body size count exceeds limit"), "Wasm validation {unit}: {message}");
            println!("parity unit={unit} chunks={} validation=shared_size_limit error={message:?}", a.len());
        }
    }
}

fn consume(chunks: Vec<Chunk>) -> (usize, usize, bool) {
    let counts = (chunks.len(), chunks.iter().map(|c| c.code.len()).sum());
    let prepared = match PreparedLinkedCode::prepare(chunks, 64 * 1024 * 1024) {
        Ok(owner) => { std::hint::black_box(&owner); drop(owner); true }
        Err(raw) => { std::hint::black_box(&raw); drop(raw); false }
    };
    (counts.0, counts.1, prepared)
}

fn main() {
    assert!(cfg!(debug_assertions), "use the debug build");
    vybe_language_php::register();
    let language = vybe_compiler::languages::find_by_name("php").unwrap();
    let options = vybe_runtime::registry::compiler_optimizations(language);
    let profile = Arc::new(vybe_runtime::profile::parse_profile((language.profile_source)()).unwrap());
    let mut bootstrap = VM::new();
    vybe_compiler::primitives::platforms::register_platforms_all(&mut bootstrap);
    let linear = ImportInterningMode::LinearControl;
    let indexed = ImportInterningMode::Indexed;
    println!("catalog helpers={} argument_emitters={} request_pool=fresh_per_compile",
        vybe_language_php::emitter::serialization_adapter::HELPER_DEFINITIONS.len(),
        vybe_language_php::emitter::argument_emission::ARGUMENT_DEFINITIONS.len());
    for (i, (source, expected)) in [
        ("<?php $a=3; $b=4; return $a+$b;", "F64(7.0)"),
        ("<?php function f($x){return $x.'!';} return f('value');", "String(\"value!\")"),
        ("<?php $a=3; $b=&$a; $b=8; return $a;", "F64(8.0)"),
        ("<?php $a=['x'=>3]; return $a['x'];", "F64(3.0)"),
        ("<?php namespace A; function f(){return 4;} namespace B; function f(){return 7;} return \\A\\f()+f();", "F64(11.0)"),
        ("<?php $x=['x'=>1]; $a=serialize($x); $b=serialize($x); return strrev('abc');", "String(\"cba\")"),
    ].iter().enumerate() {
        let ast = vybe_language_php::parse(source).unwrap();
        assert_eq!(canonical(&format!("{ast:?}")), canonical(&format!("{:?}", vybe_language_php::parse_with_peg(source).unwrap())));
        let a = compile(&ast, &profile, options, linear); let b = compile(&ast, &profile, options, indexed);
        if i == 5 {
            for chunks in [&a, &b] {
                assert_eq!(chunks.iter().filter(|chunk| chunk.name == "__php_serialize_value").count(), 1,
                    "registered same-line helper must be admitted in both lanes");
            }
            let literal = vybe_language_php::parse("<?php return strrev('abc');").unwrap();
            let selected = compile(&literal, &profile, options, indexed);
            let legacy = Compiler::with_shared_profile(profile.clone()).with_optimizations(options)
                .with_run_module_end_hooks(false).compile(&literal).unwrap();
            let bytes = |chunks: &[Chunk]| chunks.iter().map(|chunk| chunk.code.len()).sum::<usize>();
            assert!(bytes(&selected) + 1_000 < bytes(&legacy), "literal argument-emitter must be admitted");
            println!("admission serialize_bodies=1 literal_strrev_selected_bytes={} legacy_bytes={}", bytes(&selected), bytes(&legacy));
        }
        same(&a, &b, &format!("fixture-{i}"));
        for chunks in [a, b] {
            let mut vm = VM::new(); vybe_compiler::primitives::platforms::register_platforms_all(&mut vm);
            assert_eq!(format!("{:?}", vm.run(chunks).unwrap()), *expected, "VM oracle fixture {i}");
        }
    }
    for path in [
        "examples/webroot/wordpress/wp-includes/functions.php",
        "examples/webroot/wordpress/wp-includes/class-wp-query.php",
        "examples/webroot/wordpress/wp-includes/formatting.php",
        "examples/webroot/wordpress/wp-includes/script-loader.php",
    ] {
        let source = std::fs::read_to_string(path).unwrap();
        let ast = vybe_language_php::parse(&source).unwrap();
        assert_eq!(canonical(&format!("{ast:?}")), canonical(&format!("{:?}", vybe_language_php::parse_with_peg(&source).unwrap())), "full AST {path}");
        let a = compile(&ast, &profile, options, linear); let b = compile(&ast, &profile, options, indexed);
        same(&a, &b, path); drop(a); drop(b);
        for round in 0..6 {
            let indexed_first = round % 2 == 0;
            let mut inventories = Vec::new();
            for selected in if indexed_first { [true, false] } else { [false, true] } {
                let mode = if selected { indexed } else { linear };
                let wall = Instant::now(); let cpu = cpu_ns();
                let inventory = consume(compile(&ast, &profile, options, mode));
                let elapsed = cpu_ns() - cpu;
                println!("sample unit={path} round={round} indexed_first={indexed_first} indexed={selected} cpu_ns={elapsed} wall_ns={} chunks={} code={} prepared={}", wall.elapsed().as_nanos(), inventory.0, inventory.1, inventory.2);
                inventories.push(inventory);
            }
            assert_eq!(inventories[0], inventories[1]);
        }
    }
}
