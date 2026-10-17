//! Operand traffic baseline using only correctly typed core-Wasm operations.
use crate::{Chunk, Op, Value, VM};
use std::{hint::black_box, time::Instant};

fn loaded(roundtrips:usize,unsigned:bool,history:bool)->(VM,Value) {
    let mut operations=Vec::with_capacity(roundtrips*3);
    for _ in 0..roundtrips {
        operations.push(if unsigned {Op::F64_CONVERT_I32_U}else{Op::F64_FROM_I32});
        operations.push(Op::I64_TRUNC_F64_S);
        operations.push(Op::I32_WRAP_I64);
    }
    loaded_program(&operations,history)
}

fn loaded_program(operations:&[Op],history:bool)->(VM,Value) {
    let mut entry=Chunk::new("roundtrip-entry");
    entry.emit_op_u16(Op::REF_FUNC,1,1);
    entry.code.push(0); // Closure capture count is one byte.
    entry.emit_op(Op::RETURN,1);
    entry.lines.resize(entry.code.len(),1);
    let mut body=Chunk::new("roundtrip-body");
    body.arity=1;
    body.local_count=1;
    body.scratch_high_water=1;
    body.result_arity=1;
    body.emit_op_u16(Op::LOCAL_GET,0,1);
    for &operation in operations {body.emit_op(operation,1);}
    body.emit_op(Op::RETURN,1);
    body.lines.resize(body.code.len(),1);
    let mut vm=VM::new();
    vm.record_error_context(history);
    let callable=vm.run(vec![entry,body]).unwrap();
    (vm,callable)
}

#[test]
fn f32_roundtrip_exclusions_preserve_rounded_integer_results() {
    for history in [false,true] {
        for (convert,cases) in [
            (Op::F32_CONVERT_I32_S,vec![(16777217,16777216),(-16777217,-16777216),(i32::MAX,i32::MIN),(i32::MIN,i32::MIN)]),
            (Op::F32_CONVERT_I32_U,vec![(-1,0),(i32::MAX,i32::MIN),(i32::MIN,i32::MIN),(16777217,16777216)]),
        ] {
            let (mut vm,callable)=loaded_program(&[convert,Op::I64_TRUNC_F32_S,Op::I32_WRAP_I64],history);
            for (input,expected) in cases {
                assert_eq!(vm.invoke(&callable,&[Value::I32(input)]).unwrap(),Value::I32(expected));
            }
        }
    }
}

#[test]
fn scalar_truncation_fractional_boundaries_match_standard_wasm() {
    let good32=|pairs:&[(f64,i32)],single:bool| pairs.iter().map(|&(input,result)| {
        (if single {Value::F32(input as f32)}else{Value::F64(input)},Value::I32(result))
    }).collect::<Vec<_>>();
    let good64=|pairs:&[(f64,i64)]| pairs.iter().map(|&(input,result)|
        (Value::F64(input),Value::I64(result))).collect::<Vec<_>>();
    let bad=|values:&[f64],single:bool| values.iter().copied()
        .chain([f64::NAN,f64::INFINITY,f64::NEG_INFINITY])
        .map(|input| if single {Value::F32(input as f32)}else{Value::F64(input)})
        .collect::<Vec<_>>();
    let cases=[
        (Op::I32_FROM_F64,
            good32(&[(-2147483648.5,i32::MIN),(2147483647.75,i32::MAX),(-0.75,0)],false),
            bad(&[-2147483649.0,2147483648.0,4294967295.0],false)),
        (Op::I32_TRUNC_F64_U,
            good32(&[(-0.75,0),(-0.9999999999999999,0),(4294967295.5,-1)],false),
            bad(&[-1.0,4294967296.0],false)),
        (Op::I64_TRUNC_F64_S,
            good64(&[(-0.75,0),(-9223372036854775808.0,i64::MIN)]),
            bad(&[9223372036854775808.0,-9223372036854775808.0-2048.0],false)),
        (Op::I64_TRUNC_F64_U,
            good64(&[(-0.75,0),(9223372036854775808.0,i64::MIN),(18446744073709551616.0-2048.0,-2048)]),
            bad(&[-1.0,18446744073709551616.0],false)),
        (Op::I32_TRUNC_F32_S,
            good32(&[(-2147483648.0,i32::MIN),(2147483520.0,2147483520),(-0.75,0)],true),
            bad(&[-2147483904.0,2147483648.0],true)),
        (Op::I32_TRUNC_F32_U,
            good32(&[(-0.75,0),(4294967040.0,-256)],true),
            bad(&[-1.0,4294967296.0],true)),
    ];
    for history in [false,true] {
        for (op,valid,invalid) in &cases {
            for (input,expected) in valid {
                let (mut vm,callable)=loaded_program(&[*op],history);
                assert_eq!(vm.invoke(&callable,&[input.clone()]).unwrap(),*expected,
                    "op={op:?} input={input:?} history={history}");
            }
            for input in invalid {
                let (mut vm,callable)=loaded_program(&[*op],history);
                assert!(vm.invoke(&callable,&[input.clone()]).is_err(),
                    "op={op:?} input={input:?} history={history}");
            }
        }
    }
}

#[test]
fn general_f64_i64_wrap_retains_nonidentity_and_trap_boundaries() {
    let operations=[Op::I64_TRUNC_F64_S,Op::I32_WRAP_I64];
    for history in [false,true] {
        for (input,expected) in [(0.0,0),(-0.0,0),(1.9,1),(-1.9,-1),
            (4294967295.0,-1),(4294967296.0,0),(4294967297.0,1),(-4294967297.0,-1),
            (-9223372036854775808.0,0)] {
            let (mut vm,callable)=loaded_program(&operations,history);
            assert_eq!(vm.invoke(&callable,&[Value::F64(input)]).unwrap(),Value::I32(expected));
        }
        for input in [f64::NAN,f64::INFINITY,f64::NEG_INFINITY,
            9223372036854775808.0,-9223372036854775808.0-2048.0] {
            let (mut vm,callable)=loaded_program(&operations,history);
            assert!(vm.invoke(&callable,&[Value::F64(input)]).is_err(),"input={input}");
        }
    }
}

#[test]
fn valid_i32_f64_i64_i32_roundtrips_preserve_all_boundary_bits() {
    let mut seed=0x9e37_79b9u32;
    let mut values=vec![i32::MIN,i32::MAX,-1,0,1,16777215,16777216,16777217,-16777217];
    for _ in 0..256 {
        seed^=seed<<13;seed^=seed>>17;seed^=seed<<5;
        values.push(seed as i32);
    }
    for unsigned in [false,true] {
        for history in [false,true] {
            for count in [0usize,1,16,128] {
                let (mut vm,callable)=loaded(count,unsigned,history);
                for &input in &values {
                    assert_eq!(vm.invoke(&callable,&[Value::I32(input)]).unwrap(),Value::I32(input),
                        "unsigned={unsigned} history={history} count={count} input={input}");
                }
            }
        }
    }
}

#[test]
#[ignore="loaded operand conversion/local-return traffic baseline"]
fn loaded_roundtrip_operand_traffic_benchmark() {
    for unsigned in [false,true] {
        let mut rows:Vec<_>=[0usize,1,16,128].into_iter().map(|count| {
            let (vm,callable)=loaded(count,unsigned,true);
            (count,vm,callable,Vec::<f64>::new())
        }).collect();
        let args=[Value::I32(-16777217)];
        for (_,vm,callable,_) in &mut rows {
            for _ in 0..32 {assert_eq!(vm.invoke(callable,&args).unwrap(),args[0]);}
        }
        for trial in 0..21 {
            for order in 0..4 {
                let index=(trial+order)%4;
                let (_,vm,callable,samples)=&mut rows[index];
                let start=Instant::now();
                for _ in 0..512 {
                    let output=vm.invoke(black_box(callable),black_box(&args)).unwrap();
                    assert_eq!(black_box(output),args[0]);
                }
                samples.push(start.elapsed().as_nanos() as f64/512.0);
            }
        }
        for (count,_,_,samples) in &mut rows {
            samples.sort_by(f64::total_cmp);
            println!("roundtrip-traffic unsigned={unsigned} triples={count} median_call_ns={:.2}",samples[10]);
        }
    }
}
