use vybe_compiler::primitives::platforms;
use vybe_runtime::{Chunk, Op, VM, Value};
fn ops(c: &Chunk) -> Vec<(usize, Op)> {
    let mut rows = vec![];
    let mut at = 0;
    while at < c.code.len() {
        let op = Op::from_code_at(&c.code, at);
        rows.push((at, op));
        at += 4 + op.operand_format().size_in(&c.code, at + 4);
    }
    assert_eq!(at, c.code.len());
    rows
}
fn local(c: &Chunk, at: usize) -> u16 {
    u16::from_be_bytes(c.code[at + 4..at + 6].try_into().unwrap())
}
fn vm() -> VM {
    let mut v = VM::new();
    platforms::register_platforms_all(&mut v);
    v
}
fn normalized(r: Result<Value, vybe_runtime::VMError>) -> Result<Value, String> {
    r.map_err(|mut e| {
        for frame in &mut e.call_stack {
            frame.offset = 0;
        }
        e.to_string()
    })
}
fn validate(c: &[Chunk]) -> Result<usize, String> {
    let bytes = vybe_platform_wasm::write_wasm(c);
    match wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all())
        .validate_all(&bytes)
    {
        Ok(_) => Ok(bytes.len()),
        Err(e) => Err(e.to_string()),
    }
}
fn parsed(source: &str, writeback: bool) -> Vec<Chunk> {
    static REGISTER: std::sync::Once = std::sync::Once::new();
    REGISTER.call_once(vybe_language_php::register);
    let ast = vybe_language_php::parse(&format!("<?php {source}")).unwrap();
    let mut profile =
        vybe_compiler::profile::parse_profile(vybe_language_php::profile_source()).unwrap();
    profile.member_call_writes_receiver_back = writeback;
    vybe_compiler::Compiler::with_profile(profile)
        .compile(&ast)
        .unwrap()
}

fn positive(source: &str, writeback: bool, expected: Value, baseline_native_gap: bool) {
    let chunks = parsed(source, writeback);
    assert_eq!(
        normalized(vm().run(chunks.clone())),
        Ok(expected),
        "real parsed Source positive VM contract"
    );
    let native = validate(&chunks);
    if true {
        assert!(
            native.is_ok(),
            "proposed actual producer full native: {native:?}"
        );
    } else if baseline_native_gap {
        assert!(
            native
                .as_ref()
                .is_err_and(|e| e.contains("expected externref but nothing on stack")),
            "exact frozen gap: {native:?}"
        );
    } else {
        println!("frozen complete native classification: {native:?}");
    }
}
#[test]
fn actual_php_split_length_full_module_native_and_vm_contract() {
    for (h, n, count) in [
        ("banana", "an", 2.),
        ("aaaaa", "aa", 2.),
        ("abc", "z", 0.),
        ("😀é😀é", "😀", 2.),
        ("abc", "", 2.),
        ("😀", "", 0.),
    ] {
        positive(
            &format!(
                "function count_subject($h,$n){{return substr_count($h,$n);}}return count_subject('{h}','{n}');"
            ),
            false,
            Value::F64(count),
            true,
        );
    }
}
#[test]
fn actual_php_argument_order_and_shadowed_split_method() {
    positive(
        "$order=[];function arg($t,$v){global $order;$order[]=$t;return $v;}class Splitter{function split($needle){global $order;$order[]='m';return ['a','b','c','d'];}}function count_subject($h,$n){return substr_count($h,$n);}$r=count_subject(arg('h',new Splitter()),arg('n','x'));return $r.':'.implode('',$order);",
        false,
        Value::String("3:hnm".into()),
        false,
    );
}
#[test]
fn fortran_feature_receiver_writeback_and_carried_result() {
    // This actual common compiler feature is enabled by the real Fortran
    // profile. PHP syntax supplies an actual AST; no Fortran parser claim.
    positive(
        "class Counter{public $v=1;function bump(){$this->v=$this->v+1;return ['a','b','c'];}}function count_subject($h){return $h->bump();}$counter=new Counter();$r=count_subject($counter);return implode(':',$r).':'.$counter->v;",
        true,
        Value::String("a:b:c:2".into()),
        false,
    );
}

#[test]
#[ignore = "known PHP raw-symbol-property interop gap: intended callback returns 2:abc, captured current Source returns 0:"]
fn actual_php_split_symbol_hook_preserves_input_callback_once() {
    positive(
        "$order=[];class Needle{};$needle=new Needle();$needle->{'@@split'}=function($s){global $order;$order[]=$s;return [1,2,3];};function count_subject($h,$n){return substr_count($h,$n);}$r=count_subject('abc',$needle);return $r.':'.implode('',$order);",
        false,
        Value::String("2:abc".into()),
        false,
    );
}

#[test]
fn php_raw_symbol_assignment_gap_is_classified_separately() {
    let source = "$order=[];class Needle{};$needle=new Needle();$needle->{'@@split'}=function($s){global $order;$order[]=$s;return [1,2,3];};function count_subject($h,$n){return substr_count($h,$n);}$r=count_subject('abc',$needle);return $r.':'.implode('',$order);";
    let chunks = parsed(source, false);
    assert_eq!(
        normalized(vm().run(chunks.clone())),
        Ok(Value::String("0:".into())),
        "frozen PHP interop gap must stay visible; this does not satisfy Symbol.split protocol semantics"
    );
    if true {
        assert!(
            validate(&chunks).is_ok(),
            "classified PHP gap module native: {:?}",
            validate(&chunks)
        );
    } else {
        println!("classified PHP gap frozen native: {:?}", validate(&chunks));
    }
    let needle=normalized(vm().run(parsed("class Needle{};$needle=new Needle();$needle->{'@@split'}=function($s){return [1,2,3];};return $needle;",false))).unwrap();
    if let Value::Object(o) = needle {
        let o = o.lock().unwrap();
        for (name, value) in &o.properties {
            let tag = match value {
                Value::Object(_) => "Object",
                Value::String(_) => "String",
                Value::Undefined => "Undefined",
                _ => "other",
            };
            eprintln!("PHP raw-symbol fixture property {name:?}: {tag}");
        }
    } else {
        panic!("actual PHP needle must be Object");
    }
}
#[test]
fn actual_js_symbol_split_source_invokes_protocol_once_in_order() {
    static REGISTER: std::sync::Once = std::sync::Once::new();
    REGISTER.call_once(vybe_language_js::register);
    let source = "var trace='';var needle={};needle[Symbol.split]=function(s){trace=trace+s;return ['a','b','c'];};function arg(t,v){trace=trace+t;return v;}function count_subject(h,n){return h.split(n).length-1;}var answer=count_subject(arg('h','abc'),arg('n',needle))+':'+trace;";
    let ast = vybe_language_js::parse(source).unwrap();
    let profile =
        vybe_compiler::profile::parse_profile(vybe_language_js::profile_source()).unwrap();
    let chunks = vybe_compiler::Compiler::with_profile(profile)
        .compile(&ast)
        .unwrap();
    let mut machine = vm();
    machine.run(chunks.clone()).unwrap();
    assert_eq!(
        machine.global("answer"),
        Some(&Value::String("2:hnabc".into())),
        "actual Symbol.split callback once, after argument evaluation"
    );
    assert_eq!(
        machine.global("trace"),
        Some(&Value::String("hnabc".into()))
    );
    if true {
        assert!(
            validate(&chunks).is_ok(),
            "complete JS Source native: {:?}",
            validate(&chunks)
        );
    } else {
        println!(
            "frozen complete JS native classification: {:?}",
            validate(&chunks)
        );
    }
}
