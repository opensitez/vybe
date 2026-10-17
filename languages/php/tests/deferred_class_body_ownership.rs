//! Same-AST consuming PHP compilation gate: exact emitted code/metadata/Wasm,
//! raw/prepared VM semantics, parser opt-in, allocation counts and paired CPU.
//! Run with --features service-verification -- --test-threads=1 --nocapture.
use std::sync::Arc;
use vybe_ast::Module;
use vybe_runtime::{Chunk, VM, PreparedLinkedCode};
use vybe_compiler::{Compiler, primitives::deferred_class_bodies::take_statistics};
#[path = "../examples/emission_import_interning/canonical.rs"]
mod metadata;

fn workspace() -> std::path::PathBuf { std::path::Path::new(env!("CARGO_MANIFEST_DIR")).join("../..").canonicalize().unwrap() }
fn compiler(profile: &Arc<vybe_runtime::profile::LanguageProfile>) -> Compiler {
    let language=vybe_compiler::languages::find_by_name("php").unwrap();
    Compiler::with_shared_profile(profile.clone())
        .with_optimizations(vybe_runtime::registry::compiler_optimizations(language))
        .with_helper_templates(vybe_language_php::parse,
            vybe_language_php::emitter::serialization_adapter::HELPER_DEFINITIONS,
            Arc::new(vybe_compiler::primitives::helper_templates::Pool::default()))
        .with_argument_emitters(vybe_language_php::parse,
            vybe_language_php::emitter::argument_emission::ARGUMENT_DEFINITIONS)
        .with_run_module_end_hooks(false)
}
fn compile(ast: Module, owned: bool, parser: vybe_runtime::registry::ParseFn) -> Vec<Chunk> {
    take_statistics();
    let profile = Arc::new(vybe_runtime::profile::parse_profile(vybe_language_php::profile_source()).unwrap());
    let compiler=compiler(&profile);
    let result=if owned {compiler.compile_with_imports_owned_for_validation(ast,parser)}
        else {compiler.compile_with_imports(&ast)};
    result.unwrap().chunks
}
fn same_emission(a: &[Chunk], b: &[Chunk]) -> Vec<u8> {
    assert_eq!(a.len(),b.len());
    for (a,b) in a.iter().zip(b) {
        assert_eq!(a.code,b.code); assert_eq!(a.lines,b.lines);
        assert_eq!(metadata::canonical(&format!("{a:?}")),metadata::canonical(&format!("{b:?}")));
    }
    let write=|chunks|vybe_platform_wasm::write_wasm_with_options(chunks,
        vybe_platform_wasm::WasmWriteOptions{include_vybe_metadata:false,include_name_section:true});
    let a=write(a);let b=write(b);assert_eq!(a,b);a
}
fn same(a: &[Chunk], b: &[Chunk], allow_shared_size_limit: bool) {
    let a=same_emission(a,b);
    if let Err(error)=wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all()).validate_all(&a) {
        assert!(allow_shared_size_limit && error.to_string().contains("function body size count exceeds limit"), "{error}");
        println!("full-unit validation=shared-function-size-limit error={error}");
    }
}
#[derive(Clone, Copy, Debug)]
enum Expected { Number(i32), ErrorContaining(&'static str) }
fn assert_expected(result: &Result<vybe_runtime::Value, vybe_runtime::VMError>, expected: Expected) {
    match (result, expected) {
        (Ok(value), Expected::Number(number)) => {
            let actual=match value {
                vybe_runtime::Value::I32(n)=>*n as f64,
                vybe_runtime::Value::I64(n)=>*n as f64,
                vybe_runtime::Value::F64(n)=>*n,
                other=>panic!("expected numeric {number}, got {other:?}"),
            };
            assert_eq!(actual,number as f64,"known fixture value");
        },
        (Err(error), Expected::ErrorContaining(message)) => {
            assert!(error.message.contains(message),"expected source error {message:?}, got {error:?}");
        },
        (actual,expected)=>panic!("expected {expected:?}, got {actual:?}"),
    }
}
fn outcome(chunks: Vec<Chunk>, prepared: bool, expected: Expected) -> String {
    let mut vm=VM::new();vybe_compiler::primitives::platforms::register_platforms_all(&mut vm);
    vybe_runtime::framework::init_all_on_vm(&mut vm,&[&vybe_language_php::Plugin]);
    vybe_compiler::dynamic::register_dynamic_runtime_imports(&mut vm);
    let resolved=chunks[0].imports.iter().map(|i|vm.resolve_import_target(&i.module,&i.name)
        .or_else(|e|if matches!(i.module.as_str(),"env"|"*"){Ok(vybe_runtime::ImportTarget::StdlibRedirect(i.name.clone()))}else{Err(e)}))
        .collect::<Result<Vec<_>,_>>().unwrap();
    let result=if prepared {vm.run_linked_prepared(PreparedLinkedCode::prepare(chunks,64*1024*1024).ok().unwrap(),resolved)}
        else {vm.run_linked(chunks,resolved)};
    assert_expected(&result,expected);
    // Code/offsets must remain identical: no diagnostic fields are normalized.
    format!("{result:?}")
}
fn check(source: &str, expect_donation: bool, expected: Expected, parser: vybe_runtime::registry::ParseFn) {
    let ast=parser(source).unwrap();let snapshot=format!("{ast:?}");
    let old=compile(ast.clone(),false,parser);assert_eq!(take_statistics().methods,0);
    let new=compile(ast.clone(),true,parser);let admission=take_statistics();
    assert_eq!(admission.methods>0,expect_donation,"{source}: {admission:?}");
    assert_eq!(snapshot,format!("{ast:?}"));same(&old,&new,false);
    let oracle=outcome(old.clone(),false,expected);
    assert_eq!(outcome(old,true,expected),oracle);assert_eq!(outcome(new.clone(),false,expected),oracle);assert_eq!(outcome(new,true,expected),oracle);
    println!("fixture admission={admission:?} outcome={oracle}");
}
#[test]
fn actual_php_owned_emission_parity_and_admission() {
    assert!(cfg!(debug_assertions));vybe_language_php::register();
    let cases=[
        ("<?php class A{public function value(){return 7;}}return (new A)->value();",true,Expected::Number(7)),
        ("<?php namespace X; class A{public function value(){return 7;}}return (new A)->value();",true,Expected::Number(7)),
        ("<?php class A{private $x; public function __construct($x){$this->x=$x;}public function value(){return $this->x;}public static function plus($a,$b=2){return $a+$b;}}return A::plus((new A(7))->value());",true,Expected::Number(9)),
        ("<?php $g=2;class A{public function value(&$x,...$rest){global $g;$x+=$g;return $x+count($rest);}}$v=3;return (new A)->value($v,1,2);",true,Expected::Number(7)),
        ("<?php class A{public function value(){throw new Exception('donated-error');}}return (new A)->value();",true,Expected::ErrorContaining("donated-error")),
        ("<?php class A{public function value(){return 7;}public function __destruct(){return;}}return (new A)->value();",true,Expected::Number(7)),
        ("<?php trait T{public function value(){return 7;}}class A{use T;}return (new A)->value();",false,Expected::Number(7)),
    ];
    for (source,admitted,expected) in cases {check(source,admitted,expected,vybe_language_php::parse);}
    // Same source/compiler semantics; unregistered PEG parser cannot select hook.
    check(cases[0].0,false,cases[0].2,vybe_language_php::parse_with_peg);
    for path in ["examples/webroot/wordpress/wp-includes/functions.php",
        "examples/webroot/wordpress/wp-includes/class-wp-query.php",
        "examples/webroot/wordpress/wp-includes/formatting.php",
        "examples/webroot/wordpress/wp-includes/script-loader.php"] {
        let ast=vybe_language_php::parse(&std::fs::read_to_string(workspace().join(path)).unwrap()).unwrap();
        let old=compile(ast.clone(),false,vybe_language_php::parse);
        let new=compile(ast,true,vybe_language_php::parse);let admission=take_statistics();
        same(&old,&new,true);wide_census(&new,path);println!("wordpress unit={path} admission={admission:?}");
    }
}

#[test]
fn installed_reference_body_elision_keeps_complete_emission_and_default_fallback() {
    vybe_language_php::register();
    // This compile-only fixture models validated reference annotations; the
    // separate actual service fixture supplies real provider receipt coverage.
    let source="<?php class ElisionUser{public function value(){return 7;}}function unused_elision_provider(){return new SplFileInfo('/tmp/unused.php');}return (new ElisionUser)->value();";
    let mut ast=vybe_language_php::parse(source).unwrap();
    assert!(vybe_compiler::preparation::bind_installed_libraries(&mut ast,|_,_|true)>0);
    let source_snapshot=format!("{ast:?}");
    let old=compile(ast.clone(),false,vybe_language_php::parse);
    let new=compile(ast.clone(),true,vybe_language_php::parse);let stats=take_statistics();
    assert!(stats.methods>0 && stats.installed_library_methods>0,"{stats:?}");
    assert!(stats.installed_library_statements>0,"fixture must remove real body clones");
    assert_eq!(source_snapshot,format!("{ast:?}"));same(&old,&new,false);
    let oracle=outcome(old.clone(),false,Expected::Number(7));
    assert_eq!(outcome(old.clone(),true,Expected::Number(7)),oracle);
    assert_eq!(outcome(new.clone(),false,Expected::Number(7)),oracle);
    assert_eq!(outcome(new,true,Expected::Number(7)),oracle);
    let unregistered=compile(ast,true,vybe_language_php::parse_with_peg);
    assert_eq!(take_statistics().installed_library_methods,0);
    same(&old,&unregistered,false);
}


#[derive(Clone, Copy, Default, Debug)]
struct Allocations { calls: usize, bytes: usize, reallocations: usize }
thread_local! {
    static TRACK: std::cell::Cell<bool> = const { std::cell::Cell::new(false) };
    static ALLOCATIONS: std::cell::Cell<Allocations> = const { std::cell::Cell::new(Allocations { calls: 0, bytes: 0, reallocations: 0 }) };
}
struct CountingAllocator;
unsafe impl std::alloc::GlobalAlloc for CountingAllocator {
    unsafe fn alloc(&self, layout: std::alloc::Layout) -> *mut u8 {
        let pointer=unsafe { std::alloc::GlobalAlloc::alloc(&std::alloc::System,layout) };
        record_allocation(layout.size(),false);pointer
    }
    unsafe fn dealloc(&self, pointer: *mut u8, layout: std::alloc::Layout) {
        unsafe { std::alloc::GlobalAlloc::dealloc(&std::alloc::System,pointer,layout) }
    }
    unsafe fn realloc(&self, pointer: *mut u8, old: std::alloc::Layout, size: usize) -> *mut u8 {
        let pointer=unsafe { std::alloc::GlobalAlloc::realloc(&std::alloc::System,pointer,old,size) };
        record_allocation(size,true);pointer
    }
}
#[global_allocator]
static ALLOCATOR: CountingAllocator = CountingAllocator;
fn record_allocation(bytes:usize,reallocation:bool) {
    if TRACK.try_with(|active|active.get()).unwrap_or(false) {
        let _=ALLOCATIONS.try_with(|count| {let mut count_value=count.get();
            count_value.calls+=1;count_value.bytes+=bytes;count_value.reallocations+=reallocation as usize;count.set(count_value);});
    }
}
struct AllocationScope;
impl AllocationScope {
    fn start()->Self {ALLOCATIONS.with(|count|count.set(Allocations::default()));TRACK.with(|active|active.set(true));Self}
    fn finish(self)->Allocations {TRACK.with(|active|active.set(false));ALLOCATIONS.with(|count|count.get())}
}
impl Drop for AllocationScope {fn drop(&mut self){TRACK.with(|active|active.set(false));}}
fn cpu_ns()->u128 {
    let mut t=libc::timespec {tv_sec:0,tv_nsec:0};
    assert_eq!(unsafe {libc::clock_gettime(libc::CLOCK_THREAD_CPUTIME_ID,&mut t)},0);
    t.tv_sec as u128*1_000_000_000+t.tv_nsec as u128
}
fn compile_probe(ast:Module,owned:bool,profile:&Arc<vybe_runtime::profile::LanguageProfile>) {
    let compiler=compiler(profile);
    let chunks=if owned {compiler.compile_with_imports_owned_for_validation(ast,vybe_language_php::parse)}
        else {compiler.compile_with_imports(&ast)}.unwrap().chunks;
    match PreparedLinkedCode::prepare(chunks,64*1024*1024) {
        Ok(code)=>{std::hint::black_box(&code);drop(code);},
        Err(raw)=>{std::hint::black_box(&raw);drop(raw);},
    }
}
#[test]
fn actual_wordpress_allocation_and_paired_emission_cpu() {
    assert!(cfg!(debug_assertions));vybe_language_php::register();
    let profile=Arc::new(vybe_runtime::profile::parse_profile(vybe_language_php::profile_source()).unwrap());
    let mut total_methods=0;let mut old_calls=0;let mut new_calls=0;
    for path in ["examples/webroot/wordpress/wp-includes/functions.php",
        "examples/webroot/wordpress/wp-includes/class-wp-query.php",
        "examples/webroot/wordpress/wp-includes/formatting.php",
        "examples/webroot/wordpress/wp-includes/script-loader.php"] {
        let ast=vybe_language_php::parse(&std::fs::read_to_string(workspace().join(path)).unwrap()).unwrap();
        // Warm shared lookup/catalog initialization in both lanes before
        // either allocation sample. Every compile still owns a fresh Pool.
        compile_probe(ast.clone(),false,&profile);
        compile_probe(ast.clone(),true,&profile);
        take_statistics();
        // Source cloning simulates fresh parser-owned inputs and is OUTSIDE
        // the measured compiler region. Both lanes drop that input inside it.
        for owned in [false,true] {
            let input=ast.clone();take_statistics();let scope=AllocationScope::start();
            compile_probe(input,owned,&profile);let allocations=scope.finish();let admission=take_statistics();
            if owned {new_calls+=allocations.calls;total_methods+=admission.methods;}else{old_calls+=allocations.calls;}
            println!("allocation unit={path} owned={owned} counts={allocations:?} admission={admission:?}");
        }
        // Separate CPU samples disable allocation accounting in both lanes.
        // These are compiler+finalize+prepare+drop observations, not a request.
        for round in 0..3 {
            for owned in if round%2==0 {[false,true]}else{[true,false]} {
                let input=ast.clone();take_statistics();let before=cpu_ns();
                compile_probe(input,owned,&profile);let elapsed=cpu_ns()-before;
                println!("cpu unit={path} round={round} owned={owned} thread_cpu_ns={elapsed} admission={:?}",take_statistics());
            }
        }
    }
    assert!(total_methods>0,"actual units must exercise donated method bodies");
    assert!(new_calls<old_calls,"actual compiler allocation count did not improve: {new_calls} vs {old_calls}");
}

#[test]
fn php_expression_contained_class_declines_installed_metadata_elision() {
    use vybe_compiler::primitives::deferred_class_bodies::installed_metadata_admission_for_validation;
    vybe_language_php::register();
    let source="<?php $closure=function(){class SplFileInfo{public function value(){return 9;}}};return 7;";
    let mut ast=vybe_language_php::parse(source).unwrap();
    assert!(vybe_compiler::preparation::bind_installed_libraries(&mut ast,|_,_|true)>0);
    assert!(!installed_metadata_admission_for_validation(&ast.body,false),"real PHP closure class must deny schema elision");
    let old=compile(ast.clone(),false,vybe_language_php::parse);
    let new=compile(ast,true,vybe_language_php::parse);let stats=take_statistics();
    assert_eq!(stats.installed_library_methods,0);
    same(&old,&new,false);
    let oracle=outcome(old.clone(),false,Expected::Number(7));
    assert_eq!(outcome(old,true,Expected::Number(7)),oracle);
    assert_eq!(outcome(new.clone(),false,Expected::Number(7)),oracle);
    assert_eq!(outcome(new,true,Expected::Number(7)),oracle);
}

#[test]
fn actual_warmed_include_library_body_elision_cpu_and_exact_parity() {
    use vybe_compiler::primitives::deferred_class_bodies::{
        with_installed_body_baseline_for_validation,installed_metadata_admission_for_validation,
    };
    assert!(cfg!(debug_assertions));vybe_language_php::register();
    let profile=Arc::new(vybe_runtime::profile::parse_profile(vybe_language_php::profile_source()).unwrap());
    let mut admitted_units=0;
    for path in ["examples/webroot/wordpress/wp-includes/functions.php",
        "examples/webroot/wordpress/wp-includes/class-wp-query.php",
        "examples/webroot/wordpress/wp-includes/formatting.php",
        "examples/webroot/wordpress/wp-includes/script-loader.php"] {
        let mut ast=vybe_language_php::parse(&std::fs::read_to_string(workspace().join(path)).unwrap()).unwrap();
        // Same AST in both lanes, modelling existing CaaS after a validated
        // provider install. The service test separately requires a real install
        // and provider receipt; this test does not manufacture runtime evidence.
        let references=vybe_compiler::preparation::bind_installed_libraries(&mut ast,|_,_|true);
        assert!(references>0,"{path}: expected parser-owned library surfaces");
        let old=with_installed_body_baseline_for_validation(||compile(ast.clone(),true,vybe_language_php::parse));
        let new=compile(ast.clone(),true,vybe_language_php::parse);let stats=take_statistics();
        same(&old,&new,true);
        admitted_units+=usize::from(stats.installed_library_methods>0);
        println!("installed-library unit={path} reference_headers={references} admission={stats:?}");
        // Account for the new structural admission separately from full compiler
        // cost; no expression walks, preparation, allocation sampling or parsing.
        let before=cpu_ns();
        for _ in 0..16 {std::hint::black_box(installed_metadata_admission_for_validation(std::hint::black_box(&ast.body),profile.case_sensitive));}
        println!("installed-library-header-scan unit={path} repeats=16 thread_cpu_ns={}",cpu_ns()-before);
        for baseline in [true,false] {
            let input=ast.clone();take_statistics();let scope=AllocationScope::start();
            if baseline {with_installed_body_baseline_for_validation(||compile_probe(input,true,&profile));}
            else {compile_probe(input,true,&profile);}
            let allocations=scope.finish();
            println!("installed-library-allocation unit={path} baseline={baseline} counts={allocations:?} admission={:?}",take_statistics());
        }
        for round in 0..3 {
            for baseline in if round%2==0 {[true,false]}else{[false,true]} {
                let input=ast.clone();take_statistics();let before=cpu_ns();
                if baseline {with_installed_body_baseline_for_validation(||compile_probe(input,true,&profile));}
                else {compile_probe(input,true,&profile);}
                let elapsed=cpu_ns()-before;
                println!("installed-library-cpu unit={path} round={round} baseline={baseline} thread_cpu_ns={elapsed} admission={:?}",take_statistics());
            }
        }
    }
    assert!(admitted_units>0,"actual warmed CaaS units must exercise installed-body elimination");
}

// Observational scan over already emitted chunks; never inside cost samples.
fn wide_census(chunks:&[Chunk],unit:&str) {
    let mut total=0usize;let mut removable=0usize;let mut eligible=0usize;
    for chunk in chunks {
        let old=chunk.code.len().div_ceil(4)*8;total+=old;
        let mut at=0usize;let mut needs_wide=false;
        while at<chunk.code.len() {
            assert!(at+4<=chunk.code.len(),"truncated header {}",chunk.name);
            let op=vybe_runtime::Op::decode(chunk.read_u16(at),chunk.read_u16(at+2)).expect("unknown emitted opcode");
            needs_wide |= matches!(op,vybe_runtime::Op::F64_CONST|vybe_runtime::Op::REF_FUNC);
            let width=4+op.operand_format().size_in(&chunk.code,at+4);
            at=at.checked_add(width).unwrap();assert!(at<=chunk.code.len());
        }
        if !needs_wide {eligible+=1;removable+=old;}
    }
    println!("wide unit={unit} chunks={} eligible={eligible} old_wide_bytes={total} removable_wide_bytes={removable}",chunks.len());
}

#[test]
fn actual_owned_include_worker_and_borrowed_fused_service() {
    vybe_language_php::register();let mut vm=VM::new();
    vybe_compiler::primitives::platforms::register_platforms_all(&mut vm);
    vybe_runtime::framework::init_all_on_vm(&mut vm,&[&vybe_language_php::Plugin]);
    vybe_compiler::dynamic::register_dynamic_runtime_imports(&mut vm);
    vybe_compiler::dynamic::verify_deferred_class_donation_service(vm);
}

// Exercise the real include service without a CLI build. These use the PHP
// owner's exact frozen fixtures and goldens, including caller storage scope.
fn check_current_live_scope_fixture(name: &str) {
    vybe_language_php::register();
    let language=vybe_compiler::languages::find_by_name("php").unwrap();
    let mut vm=VM::new();
    vybe_compiler::primitives::platforms::register_platforms_all(&mut vm);
    vybe_runtime::framework::init_all_on_vm(&mut vm,&[&vybe_language_php::Plugin]);
    vybe_compiler::dynamic::register_dynamic_runtime_imports(&mut vm);
    let path=workspace().join("examples/php").join(format!("{name}.php"));
    let escaped=path.to_str().unwrap().replace('\\',"\\\\").replace('\'',"\\'");
    let source=format!("<?php ob_start();require '{escaped}';return ob_get_clean();");
    let result=vybe_compiler::dynamic::RuntimeCompilerService::new(&mut vm)
        .compile_and_run_source(source,language,workspace().join("examples/php/__include_service_gate.php"))
        .unwrap();
    let expected=std::fs::read_to_string(path.with_extension("expected")).unwrap();
    match result {
        vybe_runtime::Value::String(actual)=>assert_eq!(actual.as_ref(),expected.as_str(),"{name}"),
        other=>panic!("{name}: expected captured fixture output, got {other:?}"),
    }
}

#[test]
fn current_live_scope_static_include_semantics() {
    check_current_live_scope_fixture("include_live_static_identity_probe");
}

#[test]
fn current_live_scope_foreach_source_and_byref_semantics() {
    check_current_live_scope_fixture("include_live_foreach_targets_probe");
}

#[test]
fn mismatched_compiler_profile_cannot_select_php_donation() {
    vybe_language_php::register();
    let ast=vybe_language_php::parse("<?php class DonationProfileGuard{public function value(){return 7;}}").unwrap();
    let mut profile=vybe_runtime::profile::parse_profile(vybe_language_php::profile_source()).unwrap();
    profile.name="unregistered-donation-profile".into();
    let old=Compiler::with_profile(profile.clone()).compile_with_imports(&ast).err().expect("baseline missing normalizer");
    take_statistics();
    let selected=Compiler::with_profile(profile).compile_with_imports_owned_for_validation(ast,vybe_language_php::parse)
        .err().expect("same missing normalizer");
    assert_eq!(old,selected);assert_eq!(take_statistics().methods,0);
}

#[test]
fn borrowed_prepared_library_metadata_matches_legacy_and_keeps_fallback_ast() {
    vybe_language_php::register();
    let profile=Arc::new(vybe_runtime::profile::parse_profile(vybe_language_php::profile_source()).unwrap());
    let source="<?php function borrowed_metadata_unused_provider(){return new SplFileInfo('/tmp/unused.php');}return 7;";
    let mut ast=vybe_language_php::parse(source).unwrap();
    assert!(vybe_compiler::preparation::bind_installed_libraries(&mut ast,|_,_|true)>0);
    let snapshot=format!("{ast:?}");let pointer=ast.body.as_ptr();let capacity=ast.body.capacity();
    let old=compiler(&profile).compile_with_imports(&ast).unwrap().chunks;
    take_statistics();
    let new=compiler(&profile).compile_with_imports_prepared_for_validation(&ast,vybe_language_php::parse).unwrap().chunks;
    let stats=take_statistics();
    assert_eq!(stats.methods,0,"borrowed entry must never donate source bodies");
    assert!(stats.installed_library_methods>0 && stats.installed_library_statements>0,"real body omission required: {stats:?}");
    assert_eq!(snapshot,format!("{ast:?}"));assert_eq!(ast.body.as_ptr(),pointer);assert_eq!(ast.body.capacity(),capacity);
    same(&old,&new,false);
    for prepared in [false,true] {
        assert_eq!(outcome(old.clone(),prepared,Expected::Number(7)),outcome(new.clone(),prepared,Expected::Number(7)));
    }
    take_statistics();
    let unregistered=compiler(&profile).compile_with_imports_prepared_for_validation(&ast,vybe_language_php::parse_with_peg).unwrap().chunks;
    assert_eq!(take_statistics().installed_library_methods,0);same(&old,&unregistered,false);

    let mut source_class=vybe_language_php::parse("<?php function borrowed_source_unused_provider(){return new SplFileInfo('/tmp/unused.php');}class BorrowedMetadataSource{public function value(){return 8;}}return (new BorrowedMetadataSource)->value();").unwrap();
    assert!(vybe_compiler::preparation::bind_installed_libraries(&mut source_class,|_,_|true)>0);
    let original=compiler(&profile).compile_with_imports(&source_class).unwrap().chunks;
    take_statistics();
    let fallback=compiler(&profile).compile_with_imports_prepared_for_validation(&source_class,vybe_language_php::parse).unwrap().chunks;
    let stats=take_statistics();assert_eq!(stats.methods,0);assert_eq!(stats.installed_library_methods,0);
    same(&original,&fallback,false);
    assert_eq!(outcome(fallback,false,Expected::Number(8)),outcome(original,true,Expected::Number(8)));
}

#[test]
#[ignore = "paired DEBUG compiler CPU/allocation comparison for borrowed include metadata"]
fn borrowed_include_metadata_wordpress_cpu_and_allocations() {
    assert!(cfg!(debug_assertions));vybe_language_php::register();
    let profile=Arc::new(vybe_runtime::profile::parse_profile(vybe_language_php::profile_source()).unwrap());
    let mut admitted=0;
    for path in ["examples/webroot/wordpress/wp-includes/functions.php",
        "examples/webroot/wordpress/wp-includes/class-wp-query.php",
        "examples/webroot/wordpress/wp-includes/formatting.php",
        "examples/webroot/wordpress/wp-includes/script-loader.php"] {
        let mut ast=vybe_language_php::parse(&std::fs::read_to_string(workspace().join(path)).unwrap()).unwrap();
        // Same prepared AST in both lanes, modelling a previously validated
        // provider install. The real fused-service gate checks actual receipts.
        let references=vybe_compiler::preparation::bind_installed_libraries(&mut ast,|_,_|true);
        let run=|baseline| {
            let compiler=compiler(&profile);
            if baseline {compiler.compile_with_imports(&ast)}
            else {compiler.compile_with_imports_prepared_for_validation(&ast,vybe_language_php::parse)}
                .unwrap().chunks
        };
        let old=run(true);take_statistics();let new=run(false);let stats=take_statistics();
        // Timing diagnostics must cover all units even when the unchanged
        // baseline writer rejects a unit. Exact output equality remains
        // mandatory; strict correctness tests above retain validation gates.
        let wasm=same_emission(&old,&new);
        let validation=wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all())
            .validate_all(&wasm).map(|_|()).map_err(|error|error.to_string());
        println!("borrowed-metadata-baseline-validation unit={path} result={validation:?}");
        drop(wasm);drop(old);drop(new);
        admitted+=usize::from(stats.installed_library_methods>0);
        println!("borrowed-metadata-parity unit={path} reference_headers={references} admission={stats:?}");
        for baseline in [true,false] {
            take_statistics();let scope=AllocationScope::start();
            let chunks=run(baseline);std::hint::black_box(&chunks);drop(chunks);
            let counts=scope.finish();
            println!("borrowed-metadata-allocations unit={path} baseline={baseline} counts={counts:?} admission={:?}",take_statistics());
        }
        for round in 0..5 {
            for baseline in if round%2==0 {[true,false]}else{[false,true]} {
                take_statistics();let before=cpu_ns();
                let chunks=run(baseline);std::hint::black_box(&chunks);drop(chunks);
                let elapsed=cpu_ns()-before;
                println!("borrowed-metadata-cpu unit={path} round={round} baseline={baseline} thread_cpu_ns={elapsed} admission={:?}",take_statistics());
            }
        }
    }
    assert!(admitted>0,"must exercise real installed-provider body omission on some WordPress input");
}
