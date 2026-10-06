//! Opt-in VM/host boundary diagnostic. This is not a release benchmark.

use std::time::Instant;
use vybe_runtime::{Chunk, HostContext, Op, VM, Value};

const CALLS: usize = 100_000;

fn loop_chunk(import: Option<(&str, &str)>) -> Chunk {
    let mut chunk = Chunk::new("<wasi-host-call-perf>");
    chunk.local_count = 1;
    let import = import.map(|(module, name)| chunk.add_import(module, name));
    chunk.emit_i32_const(CALLS as i32, 0);
    chunk.emit_op_u16(Op::LOCAL_SET, 0, 0);
    let outer = chunk.emit_block(0);
    let (loop_label, _) = chunk.emit_loop_s(0);
    if let Some(import) = import {
        chunk.emit_call(import, 0, 0);
        chunk.emit_op(Op::DROP, 0);
    }
    chunk.emit_op_u16(Op::LOCAL_GET, 0, 0);
    chunk.emit_i32_const(1, 0);
    chunk.emit_op(Op::I32_SUB, 0);
    chunk.emit_op_u16(Op::LOCAL_TEE, 0, 0);
    chunk.emit_br_if(0, 0);
    chunk.emit_end(0);
    chunk.patch_loop(loop_label);
    chunk.emit_end(0);
    chunk.patch_block(outer);
    chunk.emit_i32_const(0, 0);
    chunk.emit_op(Op::RETURN, 0);
    chunk
}

#[test]
#[ignore = "local host-call diagnostic; run explicitly with --ignored --nocapture"]
fn bench_host_import_boundary() {
    for (label, import) in [
        ("loop", None),
        ("noop host", Some(("bench:host", "noop"))),
        (
            "WASI clock",
            Some(("wasi:clocks/monotonic-clock", "now")),
        ),
    ] {
        let mut samples = Vec::new();
        for _ in 0..5 {
            let mut vm = VM::new();
            vybe_compiler::primitives::platforms::register_platforms(
                &mut vm,
                &vybe_runtime::capabilities::Capabilities::all(),
            );
            vm.register_host_fn(
                "bench:host",
                "noop",
                Box::new(|_ctx: &mut HostContext, _args: &[Value]| Value::Null),
            );
            let start = Instant::now();
            assert_eq!(vm.run(vec![loop_chunk(import)]).unwrap(), Value::I32(0));
            samples.push(start.elapsed().as_nanos() as f64 / CALLS as f64);
        }
        samples.sort_by(f64::total_cmp);
        println!("{label}: median {:.1} ns/call (5 debug runs)", samples[2]);
    }
}

#[test]
#[ignore = "local host registry lookup diagnostic"]
fn bench_host_registry_lookup_paths() {
    let mut vm = VM::new();
    vybe_compiler::primitives::platforms::register_platforms(
        &mut vm,
        &vybe_runtime::capabilities::Capabilities::all(),
    );
    let names = [
        ("ecma:array", "push"),
        ("ecma:math", "max"),
        ("ecma:string", "includes"),
    ];
    for &(module, name) in &names {
        assert_eq!(
            vm.host_registry
                .get(&(module.to_string(), name.to_string()))
                .copied(),
            vm.resolve_host_function_index(module, name)
        );
    }

    for (label, borrowed) in [("allocating", false), ("borrowed", true)] {
        let mut samples = Vec::new();
        for _ in 0..5 {
            let started = Instant::now();
            for _ in 0..100_000 {
                for &(module, name) in &names {
                    let module = std::hint::black_box(module);
                    let name = std::hint::black_box(name);
                    let index = if borrowed {
                        vm.resolve_host_function_index(module, name)
                    } else {
                        vm.host_registry
                            .get(&(module.to_string(), name.to_string()))
                            .copied()
                    };
                    std::hint::black_box(index);
                }
            }
            samples.push(started.elapsed());
        }
        samples.sort_unstable();
        println!("{label} host lookup: {:?} per 300k", samples[2]);
    }
}
