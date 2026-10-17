//! PRIVATE compact source-line same-AST emission/storage gate. Parse/AST comparison and output validation
//! are outside timing. Timed samples include Compiler construction, emission,
//! finalization, linked-code preparation and destruction on the current thread.
//! This is not a complete WordPress request or detached CaaS-worker benchmark.
use std::{sync::Arc, time::Instant};
use vybe_compiler::Compiler;
use vybe_runtime::{Chunk, PreparedLinkedCode, VM};

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

fn compile(module: &vybe_ast::Module, profile: &Arc<vybe_runtime::profile::LanguageProfile>, mut options: vybe_runtime::registry::CompilerOptimizations, compact: bool) -> Vec<Chunk> {
    options.compact_source_lines = compact;
    let chunks = Compiler::with_shared_profile(profile.clone()).with_optimizations(options)
        .with_helper_templates(vybe_language_php::parse,
            vybe_language_php::emitter::serialization_adapter::HELPER_DEFINITIONS,
            Arc::new(vybe_compiler::primitives::helper_templates::Pool::default()))
        .with_argument_emitters(vybe_language_php::parse,
            vybe_language_php::emitter::argument_emission::ARGUMENT_DEFINITIONS)
        .with_run_module_end_hooks(false).compile(module).unwrap();
    chunks
}

// Walk real operand widths, including closure captures and LEB immediates.
// Kept outside timed samples: IR bytes are not an instruction count.
fn instruction_census(chunks: &[Chunk]) -> (usize, std::collections::BTreeMap<String, usize>) {
    let mut count = 0;
    let mut histogram = std::collections::BTreeMap::new();
    for chunk in chunks {
        let mut at = 0;
        while at < chunk.code.len() {
            assert!(at + 4 <= chunk.code.len(), "truncated opcode {} at {at}", chunk.name);
            let op = vybe_runtime::Op::decode(chunk.read_u16(at), chunk.read_u16(at + 2))
                .expect("unknown emitted opcode");
            let width = 4 + op.operand_format().size_in(&chunk.code, at + 4);
            at = at.checked_add(width).expect("instruction offset overflow");
            assert!(at <= chunk.code.len(), "truncated operand {}", chunk.name);
            count += 1;
            *histogram.entry(format!("{op:?}")).or_insert(0) += 1;
        }
        assert_eq!(at, chunk.code.len());
    }
    (count, histogram)
}

fn same(a: &[Chunk], b: &[Chunk], unit: &str) {
    assert_eq!(a.len(), b.len(), "chunk count {unit}");
    let census = instruction_census(a);
    assert_eq!(census, instruction_census(b), "instruction census {unit}");
    let mut top: Vec<_> = census.1.into_iter().collect();
    top.sort_by(|a, b| b.1.cmp(&a.1).then(a.0.cmp(&b.0)));
    top.truncate(12);
    println!("instructions unit={unit} count={} top={top:?}", census.0);
    for (i, (a, b)) in a.iter().zip(b).enumerate() {
        assert!(a.code == b.code, "code {unit} chunk {i}");
        assert!(a.lines == b.lines, "source lines {unit} chunk {i}");
        for offset in 0..=a.code.len() {
            assert_eq!(a.get_line(offset), b.get_line(offset), "byte line {unit} chunk {i} offset {offset}");
        }
        assert_eq!(a.imports, b.imports, "imports {unit} chunk {i}");
        assert_eq!(a.global_imports, b.global_imports, "global imports {unit} chunk {i}");
        assert!(canonical(&format!("{a:?}")) == canonical(&format!("{b:?}")), "complete metadata {unit} chunk {i}");
    }
    let write = |chunks: &[Chunk]| vybe_platform_wasm::write_wasm_with_options(chunks,
        vybe_platform_wasm::WasmWriteOptions { include_vybe_metadata: false, include_name_section: true });
    if unit.starts_with("fixture-") || unit == "real-error" {
        assert!(vybe_platform_wasm::write_wasm(a) == vybe_platform_wasm::write_wasm(b),
            "complete metadata-bearing Wasm bytes {unit}");
    }
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

fn consume(chunks: Vec<Chunk>) -> (usize, usize, bool, usize) {
    let counts = (chunks.len(), chunks.iter().map(|c| c.code.len()).sum());
    let storage = chunks.iter().map(|chunk| chunk.lines.storage_bytes()).sum();
    let prepared = match PreparedLinkedCode::prepare(chunks, 64 * 1024 * 1024) {
        Ok(owner) => { std::hint::black_box(&owner); drop(owner); true }
        Err(raw) => { std::hint::black_box(&raw); drop(raw); false }
    };
    (counts.0, counts.1, prepared, storage)
}

fn linked_outcome(chunks: Vec<Chunk>, prepared: bool) -> String {
    let mut vm = VM::new();
    vybe_compiler::primitives::platforms::register_platforms_all(&mut vm);
    vybe_runtime::framework::init_all_on_vm(&mut vm, &[&vybe_language_php::Plugin]);
    vybe_compiler::dynamic::register_dynamic_runtime_imports(&mut vm);
    let resolved = chunks.first().into_iter().flat_map(|chunk| &chunk.imports)
        .map(|import| match vm.resolve_import_target(&import.module, &import.name) {
            Ok(target) => Ok(target),
            Err(_) if matches!(import.module.as_str(), "env" | "*") =>
                Ok(vybe_runtime::ImportTarget::StdlibRedirect(import.name.clone())),
            Err(error) => Err(error),
        }).collect::<Result<Vec<_>, _>>().unwrap();
    let result = if prepared {
        let code = PreparedLinkedCode::prepare(chunks, 64 * 1024 * 1024)
            .ok().expect("positive prepared admission");
        vm.run_linked_prepared(code, resolved)
    } else { vm.run_linked(chunks, resolved) };
    match result {
        Ok(value) => format!("Ok({value:?})"),
        Err(error) => format!("Err({:?},{:?},{:?})", error.message, error.to_string(), error.call_stack),
    }
}

fn main() {
    assert!(cfg!(debug_assertions), "use the debug build");
    vybe_language_php::register();
    let language = vybe_compiler::languages::find_by_name("php").unwrap();
    let options = vybe_runtime::registry::compiler_optimizations(language);
    assert!(options.compact_source_lines, "registered optimized PHP must select compact maps");
    let profile = Arc::new(vybe_runtime::profile::parse_profile((language.profile_source)()).unwrap());
    let mut bootstrap = VM::new();
    vybe_compiler::primitives::platforms::register_platforms_all(&mut bootstrap);
    let peg_options = vybe_runtime::registry::compiler_optimizations(
        vybe_runtime::registry::LanguageDef { parse: vybe_language_php::parse_with_peg, ..language });
    assert_eq!(peg_options, Default::default(), "unregistered PEG options remain default");
    let mut default_chunk = Chunk::new("default-dense");
    for _ in 0..100 { default_chunk.emit_op(vybe_runtime::Op::NOP, 3); }
    assert!(default_chunk.lines.storage_bytes() >= default_chunk.lines.len() * 4);
    let linear = false;
    let indexed = true;
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
            let raw = format!("{:?}", vm.run(chunks.clone()).unwrap());
            assert_eq!(raw, *expected, "VM oracle fixture {i}");
            let mut prepared_vm = VM::new();
            vybe_compiler::primitives::platforms::register_platforms_all(&mut prepared_vm);
            let imports = chunks.first().map(|chunk| chunk.imports.as_slice()).unwrap_or(&[]);
            let resolved = imports.iter().map(|import| match prepared_vm.resolve_import_target(&import.module, &import.name) {
                Ok(target) => Ok(target),
                Err(_) if matches!(import.module.as_str(), "env" | "*") => Ok(vybe_runtime::ImportTarget::StdlibRedirect(import.name.clone())),
                Err(error) => Err(error),
            }).collect::<Result<Vec<_>, _>>().unwrap();
            let owner = PreparedLinkedCode::prepare(chunks, 64 * 1024 * 1024).ok().expect("prepared admission");
            assert_eq!(format!("{:?}", prepared_vm.run_linked_prepared(owner, resolved).unwrap()), raw);
        }
    }
    let error_ast = vybe_language_php::parse("<?php\nfunction source_line_failure(){\n return missing_function_for_line_gate();\n}\nreturn source_line_failure();").unwrap();
    let error_dense = compile(&error_ast, &profile, options, false);
    let error_compact = compile(&error_ast, &profile, options, true);
    same(&error_dense, &error_compact, "real-error");
    let oracle = linked_outcome(error_dense.clone(), false);
    assert!(oracle.starts_with("Err("), "fixture must exercise a real runtime error");
    assert_eq!(oracle, linked_outcome(error_dense, true));
    assert_eq!(oracle, linked_outcome(error_compact.clone(), false));
    assert_eq!(oracle, linked_outcome(error_compact, true));
    println!("real-error raw_prepared_lines_and_frames=equal");
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
        assert!(a.iter().all(|chunk| chunk.lines.storage_bytes() >= chunk.lines.len() * 4),
            "explicit dense control must retain dense source maps");
        let storage = |chunks: &[Chunk]| chunks.iter().map(|chunk| chunk.lines.storage_bytes()).sum::<usize>();
        assert!(storage(&b) * 10 < storage(&a), "production compact option must be active: {path}");
        same(&a, &b, path); drop(a); drop(b);
        for round in 0..6 {
            let compact_first = round % 2 == 0;
            let mut inventories = Vec::new();
            for selected in if compact_first { [true, false] } else { [false, true] } {
                let mode = if selected { indexed } else { linear };
                let wall = Instant::now(); let cpu = cpu_ns();
                let inventory = consume(compile(&ast, &profile, options, mode));
                let elapsed = cpu_ns() - cpu;
                println!("sample unit={path} round={round} compact_first={compact_first} compact={selected} cpu_ns={elapsed} wall_ns={} chunks={} code={} prepared={} line_storage_bytes={}", wall.elapsed().as_nanos(), inventory.0, inventory.1, inventory.2, inventory.3);
                inventories.push((inventory.0,inventory.1,inventory.2));
            }
            assert_eq!(inventories[0], inventories[1]);
        }
    }
}
