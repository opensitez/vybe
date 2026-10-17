//! Both policies use one actual Compiler/Runtime/JS graph. Registration stays
//! disabled until exact output and whole-compiler CPU gates have passed.
use vybe_language_js as js;
use vybe_platform_wasm as wasm;
use vybe_compiler as compiler;
use vybe_runtime::{VM, capabilities::Capabilities};
use vybe_runtime::chunk::Chunk;

struct Output { result: compiler::primitives::CompileResult, wasm: Vec<u8>, line_bytes:usize, line_bytes_before_wasm:usize, code_bytes:usize }

fn assert_debug_exact(left: &impl std::fmt::Debug, right: &impl std::fmt::Debug, label: &str) {
    let left=format!("{left:#?}");let right=format!("{right:#?}");
    if left!=right {
        let at=left.bytes().zip(right.bytes()).position(|(a,b)|a!=b)
            .unwrap_or(left.len().min(right.len()));
        let window=|text:&str|String::from_utf8_lossy(
            &text.as_bytes()[at.saturating_sub(80).min(text.len())..(at+160).min(text.len())]).into_owned();
        panic!("{label}: first difference at {at}, lengths {}/{}; left {:?}; right {:?}",
            left.len(),right.len(),window(&left),window(&right));
    }
}

fn assert_chunk_exact(left: &Chunk, right: &Chunk, index: usize) {
    // Hash iteration order is not metadata. Check every such field exactly,
    // then compare all remaining public/private fields through complete Debug.
    assert_eq!(left.block_i32_results,right.block_i32_results);
    assert_eq!(left.i64_locals,right.i64_locals);
    assert_eq!(left.call_indirect_sigs,right.call_indirect_sigs);
    assert_eq!(left.call_indirect_canon,right.call_indirect_canon);
    assert_eq!(left.type_rec_shape,right.type_rec_shape);
    assert_eq!(left.types.len(),right.types.len());
    let mut left=left.clone();let mut right=right.clone();
    for (a,b) in left.types.iter_mut().zip(&mut right.types) {
        assert_eq!(a.field_descriptors,b.field_descriptors);
        a.field_descriptors.clear();b.field_descriptors.clear();
    }
    for chunk in [&mut left,&mut right] {
        chunk.block_i32_results.clear();chunk.i64_locals.clear();
        chunk.call_indirect_sigs.clear();chunk.call_indirect_canon.clear();
        chunk.type_rec_shape.clear();
    }
    assert_debug_exact(&left,&right,&format!("complete Chunk {index}"));
}
fn registered_options(parser: vybe_runtime::registry::ParseFn) -> vybe_runtime::registry::CompilerOptimizations {
    let language=vybe_runtime::registry::all().into_iter().find(|l|l.name=="js")
        .expect("actual JS frontend registration");
    vybe_runtime::registry::compiler_optimizations(
        vybe_runtime::registry::LanguageDef {parse:parser,..language})
}
fn compile(module: &vybe_ast::Module, fused: bool) -> Result<Output, String> {
    let mut options=registered_options(js::parse);
    options.fused_closure_analysis=fused;
    compile_options(module,options)
}
fn compile_options(module: &vybe_ast::Module, options: vybe_runtime::registry::CompilerOptimizations) -> Result<Output, String> {
    let profile=vybe_runtime::profile::parse_profile(js::profile_source())?;
    let result=compiler::Compiler::with_profile(profile).with_optimizations(options)
        .compile_with_imports(module)?;
    let line_bytes_before_wasm=result.chunks.iter().map(|c|c.lines.storage_bytes()).sum();
    let wasm=wasm::write_artifact(&result.chunks,result.component_source,
        result.component_binary.as_deref(),wasm::WasmWriteOptions {
            include_vybe_metadata:true,include_name_section:true,
        })?;
    let line_bytes=result.chunks.iter().map(|c|c.lines.storage_bytes()).sum();
    let code_bytes=result.chunks.iter().map(|c|c.code.len()).sum();
    Ok(Output {result,wasm,line_bytes,line_bytes_before_wasm,code_bytes})
}

fn assert_output_exact(a:&Output,b:&Output) {
    // Exhaustive destructuring catches future result fields. Compare all
    // unordered fields logically and all remaining fields without omissions.
    let compiler::primitives::CompileResult {chunks:ac,word_local_proofs:aw,host_imports:ah,
        component_source:acs,component_binary:acb,app_shell:aa}=&a.result;
    let compiler::primitives::CompileResult {chunks:bc,word_local_proofs:bw,host_imports:bh,
        component_source:bcs,component_binary:bcb,app_shell:ba}=&b.result;
    assert_debug_exact(&(aw,ah,acs,acb,aa),&(bw,bh,bcs,bcb,ba),"complete result metadata");
    assert_eq!(ac.len(),bc.len());
    for (index,(left,right)) in ac.iter().zip(bc).enumerate() {assert_chunk_exact(left,right,index);}
    if a.wasm!=b.wasm {
        let at=a.wasm.iter().zip(&b.wasm).position(|(a,b)|a!=b)
            .unwrap_or(a.wasm.len().min(b.wasm.len()));
        panic!("complete Wasm including metadata differs at {at}, lengths {}/{}",a.wasm.len(),b.wasm.len());
    }
    assert_eq!(a.code_bytes,b.code_bytes);
}
fn run(chunks:Vec<Chunk>)->Result<String,String> {
    let mut vm=VM::new();
    // Same actual plugin registration mechanism as the CLI, with this test
    // graph's linked plugins. No substitute hosts or Runtime implementations.
    vybe_runtime::init_all_registered(&mut vm,&Capabilities::all());
    vm.run(chunks).map(|v|format!("{v:?}")).map_err(|e|format!("{e:?}"))
}
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum Outcome { VmSuccess, ExistingVmFailure, CompileSuccess, Rejected }

fn check(source:&str,execute:bool) -> Outcome {
    assert_eq!(format!("{:?}",js::parse(source)),format!("{:?}",js::parse_peg(source)),
        "complete public/PEG Module or error {source}");
    let module = match js::parse(source) {
        Ok(module) => module,
        Err(error) => {
            assert!(!execute,"semantic fixture rejected before compilation: {error}");
            println!("original_parse_rejection {error:?}");
            return Outcome::Rejected; // Full public/PEG equality checked above.
        },
    };
    let a=compile(&module,false);let b=compile(&module,true);
    match (a,b) {
        (Ok(a),Ok(b))=>{
            assert_output_exact(&a,&b);
            assert_eq!(a.line_bytes,b.line_bytes);
            println!("exact_output code_bytes={} line_storage_bytes={} chunks={}",a.code_bytes,a.line_bytes,a.result.chunks.len());
            if execute {
                let x=run(a.result.chunks);let y=run(b.result.chunks);
                println!("vm_outcome source={source:?} original={x:?} candidate={y:?}");
                assert_eq!(x,y,"VM result {source}");
                if x.is_ok() { Outcome::VmSuccess } else {
                    println!("pre_existing_VM_failure; not successful JS semantics");
                    Outcome::ExistingVmFailure
                }
            } else { Outcome::CompileSuccess }
        }
        (Err(a),Err(b))=>{
            assert!(!execute,"semantic fixture failed before VM: {source}\n{a}");
            assert_eq!(a,b,"compile/parser error {source}");
            println!("original_compile_rejection {a:?}");
            Outcome::Rejected
        }
        _=>panic!("compile acceptance mismatch: {source}"),
    }
}
const SEMANTIC_CASES: [&str;12]=[
        "function make(){let x=1;return ()=>++x;}let f=make();if(f()!==2||f()!==3)throw 'shared capture';",
        "function make(){let x=7;return function(){return x;};}if(make()()!==7)throw 'ordinary closure';",
        "function make(x){return ()=>()=>x;}if(make(9)()()!==9)throw 'nested arrow';",
        "function make(x){return class{m(){return x;}};}let C=make(11);if(new C().m()!==11)throw 'class free name';",
        "function make(){this.x=13;return ()=>this.x;}let o={x:17,make};if(o.make()()!==13)throw 'arrow receiver';",
        "function make(){return function(){return this.x;};}let o={x:19,f:make()};if(o.f()!==19)throw 'dynamic receiver';",
        "function make(a=23,...rest){let b=a;return ()=>b+rest[0];}if(make(undefined,2)()!==25)throw 'default rest';",
        "function make({a},[b]){return ()=>a+b;}if(make({a:2},[3])()!==5)throw 'binding pattern';",
        "function make(x){try{if(x)return ()=>x;throw 3;}catch(e){return ()=>e;}}if(make(2)()!==2||make(0)()!==3)throw 'try scope';",
        "function make(x){return {m(){return x;},get y(){return x+1;}};}let o=make(3);if(o.m()!==3||o.y!==4)throw 'object method';",
        "function make(x){function g(y){return ()=>x+y;}return g(4);}if(make(3)()!==7)throw 'nested declaration';",
        "function plain(a){let b=a+1;return b;}if(plain(2)!==3)throw 'no closure canary';",
];

#[test]
fn same_owner_chunks_wasm_and_vm_preserve_closure_analysis() {
    js::register();
    let mut setup=VM::new();
    vybe_runtime::init_all_registered(&mut setup,&Capabilities::all());
    use vybe_runtime::registry::{CompilerOptimizations, ParseFn};
    assert!(!CompilerOptimizations::default().fused_closure_analysis);
    let language=vybe_runtime::registry::all().into_iter().find(|l|l.name=="js").unwrap();
    fn wrapped(source: &str) -> Result<vybe_ast::Module, String> { js::parse(source) }
    // Until activation every parser, including the ordinary registered one,
    // retains every original option. Explicit true below is test-only.
    for selected in [js::parse as ParseFn,js::parse_peg as ParseFn,
        js::parse_source_only as ParseFn,wrapped as ParseFn] {
        assert_eq!(vybe_runtime::registry::compiler_optimizations(
            vybe_runtime::registry::LanguageDef {parse:selected,..language}),
            CompilerOptimizations::default());
    }
    let cases=SEMANTIC_CASES;
    let mut failures=Vec::new();
    let mut vm_successes=0;
    let mut existing_vm_failures=0;
    let mut artifacts=0;
    // A mismatch never prevents collecting later fixture/real-file results.
    let mut audit=|label:&str,source:&str,execute:bool,require_artifact:bool| {
        let result=std::panic::catch_unwind(std::panic::AssertUnwindSafe(|| {
            let outcome=check(source,execute);
            if require_artifact {assert_ne!(outcome,Outcome::Rejected,"real source must compile");}
            outcome
        }));
        match result {
            Ok(outcome)=>{
                println!("case_outcome label={label:?} outcome={outcome:?}");
                match outcome {
                    Outcome::VmSuccess=>{vm_successes+=1;artifacts+=1;}
                    Outcome::ExistingVmFailure=>{existing_vm_failures+=1;artifacts+=1;}
                    Outcome::CompileSuccess=>artifacts+=1,
                    Outcome::Rejected=>{},
                }
            }
            Err(error)=>{
                let message=error.downcast_ref::<String>().map(String::as_str)
                    .or_else(||error.downcast_ref::<&str>().copied()).unwrap_or("non-string panic");
                failures.push(format!("{label}: {message}"));
            }
        }
    };
    for (index,source) in cases.iter().enumerate() {
        audit(&format!("semantic_{index}"),source,true,true);
    }
    for (index,source) in ["const f=(x)=>;","function f( {","f(a +);","class C{m(){return ()=>this.x;}}", "function f(){return missing + (()=>this.x)();}"].iter().enumerate() {
        audit(&format!("error_or_context_{index}"),source,false,false);
    }
    for path in ["examples/webroot/wordpress/wp-includes/js/jquery/jquery.js",
        "examples/webroot/wordpress/wp-includes/js/jquery/jquery.min.js",
        "examples/webroot/wordpress/wp-includes/js/dist/core-commands.js",
        "examples/webroot/wordpress/wp-includes/js/dist/dom-ready.js"] {
        let path=std::path::Path::new(env!("CARGO_MANIFEST_DIR")).join("../..").join(path);
        audit(&path.display().to_string(),&std::fs::read_to_string(&path).unwrap(),false,true);
    }
    drop(audit);
    println!("terminal VM_successes={vm_successes} pre_existing_VM_failures={existing_vm_failures} complete_artifacts={artifacts} differential_failures={}",failures.len());
    assert!(failures.is_empty(),"all differential failures: {failures:#?}");
}

#[cfg(unix)]
fn thread_cpu() -> std::io::Result<f64> {
    let mut time=libc::timespec {tv_sec:0,tv_nsec:0};
    if unsafe {libc::clock_gettime(libc::CLOCK_THREAD_CPUTIME_ID,&mut time)} != 0 {
        return Err(std::io::Error::last_os_error());
    }
    Ok(time.tv_sec as f64 + time.tv_nsec as f64 * 1e-9)
}

/// Explicit execution only, after the exact-output test and coordination.
/// Same compiler binary, all other options identical; includes AST/result drop.
#[cfg(unix)]
#[test]
#[ignore = "balanced compiler CPU measurement; reserve a window before running"]
fn balanced_same_owner_cpu() { balanced_cpu(false); }

#[cfg(unix)]
#[test]
#[ignore = "independent compact-line CPU measurement; reserve a window first"]
fn balanced_compact_lines_cpu() { balanced_cpu(true); }

#[cfg(unix)]
fn balanced_cpu(compact_lane: bool) {
    js::register();
    let mut setup=VM::new();
    vybe_runtime::init_all_registered(&mut setup,&Capabilities::all());
    let language=vybe_runtime::registry::all().into_iter().find(|l|l.name=="js").unwrap();
    let baseline=vybe_runtime::registry::compiler_optimizations(language);
    assert!(!baseline.fused_closure_analysis);
    let profile=std::sync::Arc::new(vybe_runtime::profile::parse_profile(js::profile_source()).unwrap());
    let root=std::path::Path::new(env!("CARGO_MANIFEST_DIR")).join("../..");
    let mut cases=Vec::new();
    for path in ["examples/webroot/wordpress/wp-includes/js/jquery/jquery.js",
        "examples/webroot/wordpress/wp-includes/js/jquery/jquery.min.js",
        "examples/webroot/wordpress/wp-includes/js/dist/core-commands.js",
        "examples/webroot/wordpress/wp-includes/js/dist/dom-ready.js"] {
        cases.push((path.to_owned(),std::fs::read_to_string(root.join(path)).unwrap()));
    }
    cases.push(("no_closure".to_owned(),"function plain(a){let b=a+1;return b;}plain(2);".to_owned()));
    for (_,source) in &cases {
        if compact_lane {
            assert_debug_exact(&js::parse(source),&js::parse_peg(source),"compact CPU public/PEG Module");
            for parser in [js::parse as vybe_runtime::registry::ParseFn,js::parse_peg as vybe_runtime::registry::ParseFn] {
                let module=parser(source).unwrap();
                let mut options=registered_options(parser);
                options.fused_closure_analysis=false;
                options.compact_source_lines=false;
                let dense=compile_options(&module,options).unwrap();
                options.compact_source_lines=true;
                let compact=compile_options(&module,options).unwrap();
                assert_output_exact(&dense,&compact);
            }
        } else {
            assert_eq!(check(source,false),Outcome::CompileSuccess);
        }
    }
    println!("lane={}",if compact_lane {"compact_lines"} else {"fused_closure"});
    println!("round,input,parser,selected,source_bytes,source_lines,parse_cpu,compile_and_drop_cpu,whole_cpu");
    for round in 0..6 {
        for (index,(label,source)) in cases.iter().enumerate() {
            let order=if (round+index)%2==0 {[false,true]} else {[true,false]};
            for (mode,parser) in [("public",js::parse as vybe_runtime::registry::ParseFn),
                ("peg",js::parse_peg as vybe_runtime::registry::ParseFn)] {
                for fused in order {
                    let start=thread_cpu().unwrap();
                    let module=parser(std::hint::black_box(source)).unwrap();
                    let parsed=thread_cpu().unwrap();
                    let mut options=registered_options(parser);
                    if compact_lane { options.fused_closure_analysis=false;options.compact_source_lines=fused; }
                    else { options.fused_closure_analysis=fused; }
                    let output=compiler::Compiler::with_shared_profile(profile.clone())
                        .with_optimizations(options).compile_with_imports(&module).unwrap();
                    std::hint::black_box(&output);
                    drop(output);drop(module);
                    let end=thread_cpu().unwrap();
                    println!("{round},{label},{mode},{fused},{},{},{:.9},{:.9},{:.9}",
                        source.len(),source.lines().count(),parsed-start,end-parsed,end-start);
                }
            }
        }
    }
}

/// Independent use of the already-existing line-storage option. Fusion remains
/// false in both lanes. Retained capacity is not peak RSS or allocation traffic.
#[test]
#[ignore = "separate compact-line artifact/VM/capacity gate; coordinate execution"]
fn compact_line_artifacts_and_retained_capacity() {
    js::register();
    let mut setup=VM::new();
    vybe_runtime::init_all_registered(&mut setup,&Capabilities::all());
    assert!(!registered_options(js::parse).compact_source_lines,
        "registration must remain disabled until this gate and CPU comparison pass");
    let mut failures=Vec::new();
    let mut vm_successes=0;
    let mut baseline_failures=0;
    let mut real_totals=std::collections::BTreeMap::<String,(usize,usize,usize,usize,usize)>::new();
    let mut audit=|label:&str,mode:&str,parser:vybe_runtime::registry::ParseFn,source:&str,execute:bool,real:bool| {
        let result=std::panic::catch_unwind(std::panic::AssertUnwindSafe(|| {
            assert_debug_exact(&js::parse(source),&js::parse_peg(source),"public/PEG complete Module/error");
            let module=match parser(source) {
                Ok(module)=>module,
                Err(error)=>{
                    assert!(!execute&&!real,"required source rejected: {error}");
                    println!("compact_original_parse_error label={label:?} error={error:?}");
                    return;
                }
            };
            let mut options=registered_options(parser);
            options.fused_closure_analysis=false;
            options.compact_source_lines=false;
            let dense=compile_options(&module,options);
            options.compact_source_lines=true;
            let compact=compile_options(&module,options);
            let (dense,compact)=match (dense,compact) {
                (Ok(a),Ok(b))=>(a,b),
                (Err(a),Err(b))=>{
                    assert_eq!(a,b,"compact compile error parity");
                    assert!(!execute&&!real,"required source failed compilation: {a}");
                    println!("compact_original_compile_error label={label:?} error={a:?}");
                    return;
                }
                _=>panic!("compact compile acceptance mismatch {label}"),
            };
            assert_output_exact(&dense,&compact);
            let entries:usize=dense.result.chunks.iter().map(|c|c.lines.len()).sum();
            assert_eq!(entries,compact.result.chunks.iter().map(|c|c.lines.len()).sum::<usize>());
            let dense_after_audit:usize=dense.result.chunks.iter().map(|c|c.lines.storage_bytes()).sum();
            let compact_after_audit:usize=compact.result.chunks.iter().map(|c|c.lines.storage_bytes()).sum();
            println!("compact_census label={label:?} parser={mode} source_bytes={} source_lines={} code_bytes={} logical_line_entries={entries} dense_before_wasm={} dense_after_wasm={} dense_after_audit={dense_after_audit} compact_before_wasm={} compact_after_wasm={} compact_after_audit={compact_after_audit}",
                source.len(),source.lines().count(),dense.code_bytes,dense.line_bytes_before_wasm,dense.line_bytes,
                compact.line_bytes_before_wasm,compact.line_bytes);
            if real {
                let totals=real_totals.entry(mode.to_owned()).or_default();
                totals.0+=dense.line_bytes;totals.1+=compact.line_bytes;totals.2+=source.len();totals.3+=dense.code_bytes;totals.4+=1;
            }
            if execute {
                let a=run(dense.result.chunks);let b=run(compact.result.chunks);
                println!("compact_vm_outcome label={label:?} original={a:?} candidate={b:?}");
                assert_eq!(a,b,"compact VM result/error parity");
                if a.is_ok() {vm_successes+=1;} else {baseline_failures+=1;}
            }
        }));
        if let Err(error)=result {
            let message=error.downcast_ref::<String>().map(String::as_str)
                .or_else(||error.downcast_ref::<&str>().copied()).unwrap_or("non-string panic");
            failures.push(format!("{label}/{mode}: {message}"));
        }
    };
    for (index,source) in SEMANTIC_CASES.iter().enumerate() {
        audit(&format!("semantic_{index}"),"public",js::parse,source,true,false);
    }
    for (index,source) in ["const f=(x)=>;","function f( {","f(a +);",
        "class C{m(){return ()=>this.x;}}","function f(){return missing + (()=>this.x)();}"].iter().enumerate() {
        audit(&format!("error_or_context_{index}"),"public",js::parse,source,false,false);
    }
    let root=std::path::Path::new(env!("CARGO_MANIFEST_DIR")).join("../..");
    for path in ["examples/webroot/wordpress/wp-includes/js/jquery/jquery.js",
        "examples/webroot/wordpress/wp-includes/js/jquery/jquery.min.js",
        "examples/webroot/wordpress/wp-includes/js/dist/core-commands.js",
        "examples/webroot/wordpress/wp-includes/js/dist/dom-ready.js"] {
        let source=std::fs::read_to_string(root.join(path)).unwrap();
        for (mode,parser) in [("public",js::parse as vybe_runtime::registry::ParseFn),
            ("peg",js::parse_peg as vybe_runtime::registry::ParseFn)] {
            audit(path,mode,parser,&source,false,true);
        }
    }
    drop(audit);
    assert_eq!(real_totals.len(),2,"both parser routes must complete real-input artifact checks");
    for (mode,(dense,compact,source_bytes,code_bytes,inputs)) in real_totals {
        println!("compact_real_total parser={mode} inputs={inputs} source_bytes={source_bytes} code_bytes={code_bytes} dense_line_bytes={dense} compact_line_bytes={compact}");
        assert_eq!(inputs,4,"every real input must pass for {mode}");
    }
    println!("compact_terminal VM_successes={vm_successes} pre_existing_VM_failures={baseline_failures} differential_failures={}",failures.len());
    assert!(failures.is_empty(),"all compact-line differential failures: {failures:#?}");
}
