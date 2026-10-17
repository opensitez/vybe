use crate::{Chunk, Op, Value, VM};
fn function(second: bool) -> Vec<Chunk> {
    let mut root = Chunk::new("iterator-owner");
    root.emit_op_u16(Op::REF_FUNC, 1, 0);
    root.emit(0, 0);
    root.emit_op(Op::RETURN, 0);
    let mut c = Chunk::new("iterator-next-after-advance");
    c.arity = 2; c.param_count = 2; c.local_count = 3;
    c.emit_op_u16(Op::LOCAL_GET, 0, 0);
    c.emit_op(Op::STRING_AS_ITER, 0);
    c.emit_op_u16(Op::LOCAL_TEE, 2, 0);
    c.emit_op_u16(Op::LOCAL_GET, 1, 0);
    c.emit_op(Op::STRINGVIEW_ITER_ADVANCE, 0);
    c.emit_op(Op::DROP, 0);
    if second {
        c.emit_op_u16(Op::LOCAL_GET, 2, 0);
        c.emit_op(Op::STRINGVIEW_ITER_NEXT, 0);
        c.emit_op(Op::DROP, 0);
    }
    c.emit_op_u16(Op::LOCAL_GET, 2, 0);
    c.emit_op(Op::STRINGVIEW_ITER_NEXT, 0);
    c.emit_op(Op::RETURN, 0);
    vec![root, c]
}
#[test]
fn unicode_positions_exhaustion_and_advance_match_scalar_iterator() {
    let mut workers = [false, true].map(|second| {
        let mut vm=VM::new();let f=vm.run(function(second)).unwrap();(vm,f)
    });
    let mut texts: Vec<String> = ["", "a\0b", "é中😀", "😀a😀", "A\0Ωé中😀𐐀"].into_iter().map(str::to_string).collect();
    for base in (0..=0x10ffff).step_by(1024) {
        texts.push((base..=(base+1023).min(0x10ffff)).filter_map(char::from_u32).collect());
    }
    let mut checks=0;
    for text in texts {
        let chars: Vec<_> = text.chars().collect();
        for pos in [0,1,2,chars.len()/2,chars.len(),chars.len()+1,u32::MAX as usize] {
            for (second,(vm,f)) in workers.iter_mut().enumerate() {
                let at = pos.min(chars.len()).saturating_add(second);
                let expected = chars.get(at).map_or(-1,|cp|*cp as i32);
                assert_eq!(vm.invoke(f,&[Value::String(text.clone().into()),Value::I32(pos as u32 as i32)]).unwrap(),Value::I32(expected));checks+=1;
            }
        }
    }
    eprintln!("stringref iterator checks={checks}");
}

#[test]
#[ignore]
fn debug_sample() {
    use std::{hint::black_box,time::Instant};
    let name=std::env::var("VYBE_ITER_CASE").unwrap();
    let text=match name.as_str(){"empty"=>String::new(),"tiny"=>"a😀".into(),"ascii64k"=>"a".repeat(65536),"ascii1m"=>"a".repeat(1048576),"unicode1m"=>"Ωé中😀".repeat(1048576/11),_=>panic!("unknown case")};
    let expected=text.chars().next().map_or(-1,|cp|cp as i32);
    let mut vm=VM::new();let mut create=Chunk::new("create-view");create.emit_string_const(&text,0);create.emit_op(Op::STRING_AS_ITER,0);create.emit_op(Op::RETURN,0);let view=vm.run(vec![create]).unwrap();
    let mut cs=function(false);let c=&mut cs[1];c.arity=1;c.param_count=1;c.local_count=1;c.code.clear();
    c.emit_op_u16(Op::LOCAL_GET,0,0);c.emit_i32_const(1,0);c.emit_op(Op::STRINGVIEW_ITER_REWIND,0);c.emit_op(Op::DROP,0);c.emit_op_u16(Op::LOCAL_GET,0,0);c.emit_op(Op::STRINGVIEW_ITER_NEXT,0);c.emit_op(Op::RETURN,0);
    let f=vm.run(cs).unwrap();let args=[view];for _ in 0..8{assert_eq!(vm.invoke(&f,&args).unwrap(),Value::I32(expected));}
    let start=Instant::now();let cpu=thread_cpu();let mut n=0;while start.elapsed().as_secs_f64()<0.1{assert_eq!(black_box(vm.invoke(&f,&args).unwrap()),Value::I32(expected));n+=1;}let secs=thread_cpu()-cpu;
    println!("ITER_SAMPLE {name} {} {} {n}",secs/n as f64,start.elapsed().as_secs_f64()/n as f64);
}
#[cfg(target_os="macos")]
fn thread_cpu()->f64{#[repr(C)]struct T{sec:i64,nsec:i64}unsafe extern "C"{fn clock_gettime(c:i32,t:*mut T)->i32;}let mut t=T{sec:0,nsec:0};assert_eq!(unsafe{clock_gettime(16,&mut t)},0);t.sec as f64+t.nsec as f64*1e-9}
#[cfg(not(target_os="macos"))]
fn thread_cpu()->f64 { panic!("this ignored timing probe requires macOS thread CPU time") }
