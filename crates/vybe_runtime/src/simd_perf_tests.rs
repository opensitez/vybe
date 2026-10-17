use super::*;
use std::hint::black_box;
use std::time::Instant;

fn previous_i32_add(vm: &mut VM) -> Result<(), VMError> {
    let (a, b) = vm.pop_v128_pair_fast().unwrap();
    let mut out = [0u8; 16];
    for lane in 0..4 {
        write_i32_lane(&mut out, lane, lane_i32(&a, lane).wrapping_add(lane_i32(&b, lane)));
    }
    vm.push_fast(Value::V128(out));
    Ok(())
}

fn previous_i8_add(vm: &mut VM) -> Result<(), VMError> {
    let (a, b) = vm.pop_v128_pair_fast().unwrap();
    let mut out = [0u8; 16];
    for lane in 0..16 { out[lane] = a[lane].wrapping_add(b[lane]); }
    vm.push_fast(Value::V128(out));
    Ok(())
}

fn previous_f32_add(vm: &mut VM) -> Result<(), VMError> {
    let (a, b) = vm.pop_v128_pair_fast().unwrap();
    let mut out = [0u8; 16];
    for lane in 0..4 {
        write_f32_lane(&mut out, lane, lane_f32(&a, lane) + lane_f32(&b, lane));
    }
    vm.push_fast(Value::V128(out));
    Ok(())
}

fn current_f32_add(vm: &mut VM) -> Result<(), VMError> {
    vm.simd_f32x4_binop(|a, b| a + b)
}

fn exercise(vm: &mut VM, kernel: fn(&mut VM) -> Result<(), VMError>, b: [u8; 16], iterations: usize) -> [u8; 16] {
    vm.stack.clear();
    vm.stack.push(Value::V128([0; 16]));
    for _ in 0..iterations {
        vm.stack.push(Value::V128(b));
        kernel(vm).unwrap();
    }
    assert_eq!(vm.stack.len(), 1);
    let Value::V128(out) = vm.stack[0] else { panic!("vector result") };
    black_box(out)
}

#[test]
#[ignore = "paired SIMD stack/native-kernel benchmark, not VM dispatch timing"]
fn simd_kernel_benchmark() {
    type Kernel = fn(&mut VM) -> Result<(), VMError>;
    let mut float_one = [0u8; 16];
    for lane in 0..4 { write_f32_lane(&mut float_one, lane, 1.0); }
    let cases: &[(&str, Kernel, Kernel, [u8; 16])] = &[
        ("i8x16.add", previous_i8_add, VM::simd_i8x16_add, [3; 16]),
        ("i32x4.add", previous_i32_add, VM::simd_i32x4_add, [3; 16]),
        ("f32x4.add", previous_f32_add, current_f32_add, float_one),
    ];
    let mut vm = VM::new();
    vm.stack.reserve(4);
    for &(name, previous, current, b) in cases {
        let expected = exercise(&mut vm, previous, b, 1000);
        assert_eq!(exercise(&mut vm, current, b, 1000), expected);
        for sample in 0..5 {
            let mut times = [std::time::Duration::ZERO; 2];
            let mut outputs = [[0u8; 16]; 2];
            for candidate in if sample % 2 == 0 { [false, true] } else { [true, false] } {
                let kernel = if candidate { current } else { previous };
                let start = Instant::now();
                outputs[candidate as usize] = exercise(&mut vm, kernel, b, 1_000_000);
                times[candidate as usize] = start.elapsed();
            }
            assert_eq!(outputs[0], outputs[1]);
            println!("simd name={name} sample={sample} previous={:?} current={:?} ratio={:.3}",
                times[0], times[1], times[0].as_secs_f64() / times[1].as_secs_f64());
        }
    }
}

