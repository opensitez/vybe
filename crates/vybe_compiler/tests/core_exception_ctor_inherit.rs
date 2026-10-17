#![cfg(feature = "primitive-benchmarks")]
use vybe_ast::{ClassLinkage, ClassMember, StmtKind};
use vybe_compiler::primitives;
use vybe_runtime::{Chunk, VM};
fn compile(source: &str, fast: bool) -> Result<Vec<Chunk>, String> {
    let inherit_fields=fast;
    static REGISTER: std::sync::Once = std::sync::Once::new();
    REGISTER.call_once(vybe_language_php::register);
    let mut module = vybe_language_php::parse(source).unwrap();
    if !inherit_fields {
        let fields:Vec<_>=module.body.iter().find_map(|st| {
            if let StmtKind::ClassDecl {name,members,..}=&st.kind {
                if name=="Exception" {return Some(members.iter().filter(|m|matches!(m,ClassMember::Field{name,..} if ["message","code","previous","cause"].contains(&name.as_str()))).cloned().collect::<Vec<_>>());}
            }None
        }).unwrap();
        assert_eq!(fields.len(),4);
        for st in &mut module.body {
            if let StmtKind::ClassDecl {parents,members,modifiers,..}=&mut st.kind {
                if let ClassLinkage::Library {identity}=&mut modifiers.linkage {
                    if identity.starts_with("php:core-exceptions:v5:") {
                        *identity=identity.replace(":v5:",":v2:");
                        if !parents.is_empty() {members.splice(0..0,fields.clone());}
                    }
                }
            }
        }
    }
    let profile =
        vybe_compiler::profile::parse_profile(vybe_language_php::profile_source()).unwrap();
    vybe_compiler::Compiler::with_profile(profile).compile(&module)
}
fn vm() -> VM {
    let mut v = VM::new();
    primitives::platforms::register_platforms_all(&mut v);
    v
}
fn output(n: &str, v: &serde_json::Value) {
    std::fs::write(
        std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
            .join(format!("../../CORE_EXCEPTION_CTOR_{n}.json")),
        serde_json::to_vec_pretty(v).unwrap(),
    )
    .unwrap();
}
#[test]
fn builtin_exception_fields_preserve_behavior_and_reduce_code() {
    let mut rows = vec![];
    for (label, source) in [
        (
            "error",
            "<?php $e=new TypeError('message',7); return get_class($e).':'.$e->getMessage().':'.$e->getCode().':'.$e->__toString();",
        ),
        (
            "grandchild",
            "<?php $e=new DivisionByZeroError('message',7); return get_class($e).':'.$e->getMessage().':'.$e->getCode().':'.$e->__toString();",
        ),
        (
            "exception",
            "<?php $e=new PDOException('message',7); return get_class($e).':'.$e->getMessage().':'.$e->getCode().':'.$e->__toString();",
        ),
        (
            "default",
            "<?php $e=new TypeError(); return $e->getMessage().':'.$e->getCode();",
        ),
        (
            "null",
            "<?php $e=new TypeError(null,null,null); return $e->getMessage().':'.$e->getCode();",
        ),
        (
            "previous",
            "<?php $p=new Exception('previous'); $e=new PDOException('message',7,$p); return $e->getPrevious()===$p;",
        ),
        (
            "callback",
            "<?php $e=new TypeError('first'); $f=[$e,'__construct']; $f('second',9); return $e->getMessage().':'.$e->getCode();",
        ),
        (
            "named",
            "<?php $e=new TypeError(code:9,message:'named'); return $e->getMessage().':'.$e->getCode();",
        ),
        (
            "user-inherit",
            "<?php class C extends LogicException {} $e=new C('message',7); return $e->getMessage().':'.$e->getCode();",
        ),
        (
            "explicit-parent",
            "<?php class C extends LogicException { public function __construct($m){parent::__construct('p:'.$m,5);} } $e=new C('message'); return $e->getMessage().':'.$e->getCode();",
        ),
        (
            "own-constructor",
            "<?php class C extends LogicException { public function __construct(){ $this->message='own'; } } $e=new C(); return $e->getMessage().':'.$e->getCode();",
        ),
        (
            "named-default-message",
            "<?php $e=new TypeError(code:9); return $e->getMessage().':'.$e->getCode();",
        ),
        (
            "named-previous",
            "<?php $p=new Exception('p'); $e=new DivisionByZeroError(previous:$p); return $e->getPrevious()===$p;",
        ),
        (
            "error-exception",
            "<?php $e=new ErrorException('message',7); return $e->getMessage().':'.$e->getCode();",
        ),
        (
            "receiver-before-parent",
            "<?php class C extends LogicException { public $x; public function __construct(){ $this->x=11; parent::__construct('p',5); } } $e=new C(); return $e->x.':'.$e->getMessage().':'.$e->getCode();",
        ),
        (
            "throw",
            "<?php try { throw new TypeError('caught',7); } catch (Throwable $e) { return $e->getMessage().':'.$e->getCode(); }",
        ),
    ] {
        let mut sizes = vec![];
        let mut outcomes = vec![];
        for fast in [false, true] {
            match compile(source, fast) {
                Ok(cs) => {
                    sizes.push(serde_json::json!({"bytes":cs.iter().map(|c|c.code.len()).sum::<usize>(),"chunks":cs.len(),"construct_method_bytes":cs.iter().filter(|c|c.name=="__construct").map(|c|c.code.len()).sum::<usize>()}));
                    outcomes.push(
                        vm().run(cs)
                            .map(|v| format!("{v:?}"))
                            .map_err(|e| e.message),
                    );
                }
                Err(e) => {
                    sizes.push(serde_json::json!({"compile_error":e}));
                    outcomes.push(Err(e));
                }
            }
        }
        assert!(outcomes[0].is_ok(), "{label}: {outcomes:?}");
        assert_eq!(outcomes[0], outcomes[1], "{label}");
        assert!(sizes[1]["bytes"].as_u64().unwrap() < sizes[0]["bytes"].as_u64().unwrap(), "{label}: {sizes:?}");
        eprintln!("{label}: {outcomes:?} {sizes:?}");
        rows.push(
            serde_json::json!({"label":label,"source":source,"sizes":sizes,"outcomes":outcomes}),
        );
    }
    output("PROTOTYPE", &serde_json::json!({"rows":rows}));
}

#[test]
#[ignore]
fn debug_pairs() {
    use std::time::Instant;
    let mut rows=vec![];
    let compile_pair=|source| [compile(source,false).unwrap(),compile(source,true).unwrap()];
    for (label, source) in [
        ("type-error", "<?php return function(){ $e=new TypeError('message',7); return $e->getCode(); };"),
        ("grandchild", "<?php return function(){ $e=new DivisionByZeroError('message',7); return $e->getCode(); };"),
        ("user-inherit", "<?php class C extends LogicException {} return function(){ $e=new C('message',7); return $e->getCode(); };"),
    ] {
        let programs=compile_pair(source);
        // Class globals belong to a module/VM. Loading both variants into one
        // VM would make both closures resolve the last registered class!
        let mut vs=[vm(),vm()];
        let fs=[vs[0].run(programs[0].clone()).unwrap(),vs[1].run(programs[1].clone()).unwrap()];
        for side in 0..2 { for _ in 0..10 { assert_eq!(vs[side].invoke(&fs[side],&[]).unwrap().as_f64(),7.0); } }
        let mut samples=[vec![],vec![]];
        let mut cpus=[vec![],vec![]];
        for pair in 0..7 {
            let mut wall=[0.0,0.0]; let mut cpu=[0.0,0.0]; let mut n=0;
            while wall[0]<0.15 || wall[1]<0.15 || n<16 {
                for side in if (pair+n)%2==0 {[0,1]} else {[1,0]} {
                    let c=thread_cpu_seconds();let t=Instant::now();
                    std::hint::black_box(vs[side].invoke(&fs[side],&[]).unwrap());
                    wall[side]+=t.elapsed().as_secs_f64();
                    cpu[side]+=thread_cpu_seconds()-c;
                }
                n+=1;
            }
            for side in 0..2 { samples[side].push(wall[side]/n as f64);cpus[side].push(cpu[side]/n as f64); }
        }
        let mut ratios:Vec<_>=samples[0].iter().zip(&samples[1]).map(|(a,b)|a/b).collect();ratios.sort_by(f64::total_cmp);
        let mut cpu_ratios:Vec<_>=cpus[0].iter().zip(&cpus[1]).map(|(a,b)|a/b).collect();cpu_ratios.sort_by(f64::total_cmp);
        let row=serde_json::json!({"label":label,"median_speedup":ratios[3],"median_thread_cpu_speedup":cpu_ratios[3],"seconds_per_call":samples,"thread_cpu_seconds_per_call":cpus});
        eprintln!("{row}");rows.push(row);
    }
    // Fresh VMs exclude registration from timing and include module preparation
    // plus class initialization. Clone input chunks before starting the clock.
    let source="<?php $e=new TypeError('message',7); return $e->getCode();";
    let programs=compile_pair(source);
    let mut wall=[vec![],vec![]];let mut cpu=[vec![],vec![]];
    for pair in 0..15 {
        for side in if pair%2==0 {[0,1]} else {[1,0]} {
            let mut v=vm();let cs=programs[side].clone();
            let c=thread_cpu_seconds();let t=Instant::now();
            assert_eq!(v.run(cs).unwrap().as_f64(),7.0);
            wall[side].push(t.elapsed().as_secs_f64());cpu[side].push(thread_cpu_seconds()-c);
        }
    }
    let mut ratios:Vec<_>=wall[0].iter().zip(&wall[1]).map(|(a,b)|a/b).collect();ratios.sort_by(f64::total_cmp);
    let mut cpu_ratios:Vec<_>=cpu[0].iter().zip(&cpu[1]).map(|(a,b)|a/b).collect();cpu_ratios.sort_by(f64::total_cmp);
    rows.push(serde_json::json!({"label":"fresh-module","median_speedup":ratios[7],"median_thread_cpu_speedup":cpu_ratios[7],"seconds_per_call":wall,"thread_cpu_seconds_per_call":cpu}));
    output("MEASUREMENTS",&serde_json::json!({"pairs":7,"same_vm":false,"separate_class_globals":true,"cases":rows}));
}
#[cfg(target_os="macos")]
fn thread_cpu_seconds()->f64 {
    #[repr(C)] struct Timespec {sec:i64,nsec:i64}
    unsafe extern "C" {fn clock_gettime(clock:i32,result:*mut Timespec)->i32;}
    let mut result=Timespec{sec:0,nsec:0};assert_eq!(unsafe{clock_gettime(16,&mut result)},0);
    result.sec as f64+result.nsec as f64*1e-9
}
#[cfg(not(target_os="macos"))]
fn thread_cpu_seconds()->f64 {0.0}

#[test]
fn native_modules_validate() {
    let source="<?php $e=new DivisionByZeroError(code:7,message:'message'); return $e->getCode();";
    let mut rows=vec![];
    for fast in [false,true] {
        let cs=compile(source,fast).unwrap();
        let bytes=vybe_platform_wasm::write_wasm_with_options(&cs,vybe_platform_wasm::WasmWriteOptions {include_name_section:false,include_vybe_metadata:false});
        wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all()).validate_all(&bytes).unwrap();
        let root=std::path::Path::new(env!("CARGO_MANIFEST_DIR")).join("../../artifacts");
        std::fs::create_dir_all(&root).unwrap();
        std::fs::write(root.join(if fast {"core-exception-candidate.wasm"} else {"core-exception-before.wasm"}),&bytes).unwrap();
        rows.push(serde_json::json!({"fast":fast,"native_bytes":bytes.len()}));
    }
    output("NATIVE",&serde_json::json!({"rows":rows}));
}

#[test]
fn inherited_defaults_and_child_overrides() {
    let mut rows=vec![];
    for source in [
        "<?php $e=new TypeError('message',7); return $e->getMessage().':'.$e->getCode();",
        "<?php $e=new DivisionByZeroError('message',7); return $e->getMessage().':'.$e->getCode();",
        "<?php $p=new Exception('p'); $e=new PDOException('message',7,$p); return $e->getPrevious()===$p;",
        "<?php class C extends LogicException { function __construct() {} } $e=new C(); return $e->getMessage().':'.$e->getCode();",
        "<?php class C extends LogicException {public $message='own'; function __construct(){} } $e=new C(); return $e->getMessage();",
        "<?php $e=new TypeError(code:9,message:'named'); return $e->getMessage().':'.$e->getCode();",
    ] {
        let mut outcomes=vec![];let mut sizes=vec![];
        for fields in [false,true] {
            let cs=compile(source,fields).unwrap();
            sizes.push(cs.iter().map(|c|c.code.len()).sum::<usize>());
            outcomes.push(vm().run(cs).map(|v|format!("{v:?}")).map_err(|e|e.message));
        }
        assert!(outcomes[0].is_ok(),"{source}: {outcomes:?}");assert_eq!(outcomes[0],outcomes[1],"{source}");
        rows.push(serde_json::json!({"source":source,"outcomes":outcomes,"sizes":sizes}));
    }
    output("FIELDS_PROTOTYPE",&serde_json::json!({"rows":rows}));
}
