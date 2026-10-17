use vybe_compiler::primitives::{self};
use vybe_runtime::{Chunk, Op, VM, Value};
#[path = "fixtures/qp_cursor_before.rs"]
mod before;
#[path="fixtures/qp_cursor_candidate.rs"] mod candidate;
#[path = "fixtures/primitive_native_export.rs"]
mod native_export;
fn program(fast: bool) -> Vec<Chunk> {
    let mut root = Chunk::new("qp-owner");
    primitives::functions::emit_ref_func(&mut root, 1, 0, 0);
    root.emit_op(Op::RETURN, 0);
    let mut c = Chunk::new("entry");
    c.arity = 1; c.local_count = 1;
    c.emit_op_u16(Op::LOCAL_GET, 0, 0);
    let mut bodies=vec![c];
    (if fast {candidate::emit_quoted_printable_decode} else {before::emit_quoted_printable_decode})(&mut bodies,0,1,0);
    bodies[0].emit_op(Op::RETURN,0);
    let mut c=bodies.remove(0);
    c.name = "entry".into();
    c.param_count = c.arity;
    let mut cs = vec![root, c];
    primitives::bundle::finalize_with_runtime_helpers(&mut cs);
    primitives::globals::declare_free_globals(&mut cs);
    primitives::globals::normalize_global_table(&mut cs);
    cs
}
fn worker(fast: bool) -> (VM, Value) {
    let mut vm = VM::new();
    primitives::platforms::register_platforms_all(&mut vm);
    let f = vm.run(program(fast)).unwrap();
    (vm, f)
}
fn digit(u: u16) -> Option<u16> {
    match u {
        48..=57 => Some(u - 48),
        65..=70 => Some(u - 55),
        97..=102 => Some(u - 87),
        _ => None,
    }
}
fn oracle(s: &str) -> String {
    let u: Vec<_> = s.encode_utf16().collect();
    let mut out = String::new();
    let mut i = 0;
    while i < u.len() {
        if u[i] == 61 {
            let mut look = i + 1;
            while look < u.len() && (u[look] == 32 || u[look] == 9) {
                look += 1;
            }
            if look == u.len() {
                break;
            }
            if u[look] == 13 || u[look] == 10 {
                let first = u[look];
                look += 1;
                if first == 13 && look < u.len() && u[look] == 10 {
                    look += 1;
                }
                i = look;
                continue;
            }
            if look == i + 1 && i + 2 < u.len() {
                if let (Some(a), Some(b)) = (digit(u[i + 1]), digit(u[i + 2])) {
                    out.push(char::from_u32((a * 16 + b) as u32).unwrap());
                    i += 3;
                    continue;
                }
            }
        }
        out.push_str(&String::from_utf16_lossy(&[u[i]]));
        i += 1;
    }
    out
}
fn check(ws: &mut [(VM, Value); 2], s: &str) -> usize {
    let expected = Value::String(oracle(s).into());
    for (vm, f) in ws {
        assert_eq!(
            vm.invoke(f, &[Value::String(s.into())]).unwrap(),
            expected,
            "input={s:?}"
        );
    }
    2
}
#[test]
fn exhaustive_digits_and_sizes() {
    let mut ws = [worker(false), worker(true)];
    let mut checks = 0;
    let mut domain = 0;
    for base in (0..65536u32).step_by(256) {
        let mut s = String::new();
        for u in base..base + 256 {
            if let Some(ch) = char::from_u32(u) {
                s.push('=');
                s.push(ch);
                s.push_str("0|=0");
                s.push(ch);
                s.push('|');
                domain += 1;
            }
        }
        checks += check(&mut ws, &s);
    }
    let alphabet = ['=', '0', 'A', 'f', 'G', '\r', '\n', ' ', '\t', 'é'];
    for len in 0..5 {
        for mut bits in 0..10usize.pow(len) {
            let s: String = (0..len)
                .map(|_| {
                    let ch = alphabet[bits % 10];
                    bits /= 10;
                    ch
                })
                .collect();
            checks += check(&mut ws, &s);
        }
    }
    let hex = "0123456789ABCDEFabcdef";
    for a in hex.chars() {
        for b in hex.chars() {
            checks += check(&mut ws, &format!("={a}{b}"));
        }
    }
    for s in [
        "",
        "=",
        "= ",
        "=\t\r\nx",
        "abc=\rnext",
        "==41",
        "=00=FF=7f",
        "é😀=41😀é",
        "x\0=00\r\n",
        "=0",
        "=GG=G0=0G",
    ] {
        checks += check(&mut ws, s);
    }
    for input in [
        Value::Null,
        Value::Undefined,
        Value::I32(1),
        Value::Bool(true),
    ] {
        let outcomes = [false, true].map(|fast| {
            let (mut vm, f) = worker(fast);
            vm.invoke(&f, &[input.clone()])
                .unwrap_err()
                .to_string()
                .contains("TypeError")
        });
        assert_eq!(outcomes, [true, true]);
        checks += 2;
    }
    let mut sizes = vec![];
    for fast in [false, true] {
        let cs = program(fast);
        let bytes = vybe_platform_wasm::write_wasm_with_options(
            &cs,
            vybe_platform_wasm::WasmWriteOptions {
                include_name_section: false,
                include_vybe_metadata: false,
            },
        );
        wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all())
            .validate_all(&bytes)
            .unwrap();
        native_export::export_chunks(&cs, &format!("qp-cursor-{fast}"));
        sizes.push(serde_json::json!({"body":cs[1].code.len(),"locals":cs[1].local_count,"native":bytes.len()}));
    }
    assert!(sizes[1]["body"].as_u64() < sizes[0]["body"].as_u64());
    assert!(sizes[1]["native"].as_u64() < sizes[0]["native"].as_u64());
    assert_eq!(sizes[1]["locals"], sizes[0]["locals"]);
    let result =
        serde_json::json!({"checks":checks,"digit_position_scalar_domain":domain,"sizes":sizes});
    std::fs::write(
        std::path::Path::new(env!("CARGO_MANIFEST_DIR")).join("../../QP_CURSOR_GATES.json"),
        serde_json::to_vec_pretty(&result).unwrap(),
    )
    .unwrap();
    eprintln!("{result}");
}
#[test]
#[ignore]
fn debug_pairs(){
 use std::time::Instant;
 let mut rows=vec![];
 for(label,text,control)in [
  ("control","=aF".repeat(256),true),
  ("empty",String::new(),false),
  ("short","=aF".into(),false),
  ("plain256","abcd".repeat(64),false),
  ("escapes256","=aF".repeat(256),false),
  ("padding","= \t\r\n".repeat(128),false),
  ("malformed","=Gg".repeat(256),false),
 ]{
  let (va,fa)=worker(false);let(vb,fb)=worker(!control);let mut vs=[va,vb];let fs=[fa,fb];let args=[Value::String(text.clone().into())];
  for side in 0..2{for _ in 0..8{std::hint::black_box(vs[side].invoke(&fs[side],&args).unwrap());}}
  let mut wall=[vec![],vec![]];let mut cpu=[vec![],vec![]];
  for pair in 0..7{let mut secs=[0.0,0.0];let mut cpus=[0.0,0.0];let mut n=0;
   while secs[0]<0.15||secs[1]<0.15||n<8{for side in if(pair+n)%2==0{[0,1]}else{[1,0]}{let c=thread_cpu_seconds();let t=Instant::now();std::hint::black_box(vs[side].invoke(&fs[side],&args).unwrap());secs[side]+=t.elapsed().as_secs_f64();cpus[side]+=thread_cpu_seconds()-c;}n+=1;}
   for side in 0..2{wall[side].push(secs[side]/n as f64);cpu[side].push(cpus[side]/n as f64);}
  }
  let mut ratios:Vec<_>=wall[0].iter().zip(&wall[1]).map(|(a,b)|a/b).collect();ratios.sort_by(f64::total_cmp);
  let mut cr:Vec<_>=cpu[0].iter().zip(&cpu[1]).map(|(a,b)|a/b).collect();cr.sort_by(f64::total_cmp);
  rows.push(serde_json::json!({"label":label,"control":control,"median_speedup":ratios[3],"median_thread_cpu_speedup":cr[3],"seconds_per_call":wall,"thread_cpu_seconds_per_call":cpu}));
 }
 std::fs::write(std::path::Path::new(env!("CARGO_MANIFEST_DIR")).join("../../QP_CURSOR_MEASUREMENTS.json"),serde_json::to_vec_pretty(&serde_json::json!({"pairs":7,"separate_vms":true,"cases":rows})).unwrap()).unwrap();
}
#[cfg(target_os="macos")]
fn thread_cpu_seconds()->f64 {#[repr(C)]struct Timespec{sec:i64,nsec:i64}unsafe extern "C"{fn clock_gettime(clock:i32,result:*mut Timespec)->i32;}let mut t=Timespec{sec:0,nsec:0};assert_eq!(unsafe{clock_gettime(16,&mut t)},0);t.sec as f64+t.nsec as f64*1e-9}
#[cfg(not(target_os="macos"))]
fn thread_cpu_seconds()->f64 {0.0}

#[test]
fn production_matches_measured_candidate(){
 let mut root=Chunk::new("qp-owner");primitives::functions::emit_ref_func(&mut root,1,0,0);root.emit_op(Op::RETURN,0);
 let mut body=primitives::polyfills::build_runtime_helpers_for_exports(&mut root,&["__stdlib_encoding_quoted_printable_decode"]).chunks.remove(0);
 body.name="entry".into();body.param_count=body.arity;
 let mut cs=vec![root,body];primitives::bundle::finalize_with_runtime_helpers(&mut cs);primitives::globals::declare_free_globals(&mut cs);primitives::globals::normalize_global_table(&mut cs);
 let actual=cs.remove(1);let measured=program(true).remove(1);
 assert_eq!(actual.code,measured.code);assert_eq!(actual.constants,measured.constants);assert_eq!(actual.imports,measured.imports);assert_eq!(actual.local_count,measured.local_count);
}
#[test]
fn signed_cursor_bounds_are_safe(){
 for len in [0i64,1,2,3,2147483645,2147483646,2147483647]{
  for i in [0,1,len-3,len-2,len-1].into_iter().filter(|i|*i>=0&&*i<len){
   assert_eq!(i+2<len,len-i>2);assert!(i+1<=i32::MAX as i64);
   if len-i>2{assert!(i+2<=i32::MAX as i64);assert!(i+3<=len);}
   for look in [i+1,len-1,len].into_iter().filter(|l|*l>=i&&*l<=len){assert!(look-i<=len);if look<len{assert!(look+1<=len);}}
  }
 }
}
