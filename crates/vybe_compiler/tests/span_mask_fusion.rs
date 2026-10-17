use vybe_compiler::primitives::{self, memory_span, memory_span_set};
use vybe_runtime::{Chunk, Op, VM, Value};
#[path = "fixtures/span_set_mask_before.rs"]
mod before_set;
#[path = "fixtures/span_mask_before.rs"]
mod before_single;
#[path = "fixtures/primitive_native_export.rs"]
mod native_export;
fn program(fast: bool, take: bool, set: &[u8], single: bool) -> Vec<Chunk> {
    let mut root = Chunk::new("span-owner");
    root.memory_min_pages = vec![1];
    root.memory_is_64 = vec![false];
    primitives::functions::emit_ref_func(&mut root, 1, 0, 0);
    root.emit_op(Op::RETURN, 0);
    let mut c = Chunk::new("span");
    c.arity = if single { 2 } else { 1 };
    c.param_count = c.arity;
    c.local_count = c.arity as u16;
    for slot in 0..c.arity {
        c.emit_op_u16(Op::LOCAL_GET, slot as u16, 0);
    }
    let mut chunks = vec![root, c];
    if single {
        if fast {
            memory_span::emit_single_byte(&mut chunks, 1, take, 0);
        } else {
            before_single::emit_single_byte(&mut chunks, 1, take, 0);
        }
    } else if fast {
        memory_span_set::emit_literal_set(&mut chunks, 1, take, set, 0);
    } else {
        before_set::emit_literal_set(&mut chunks, 1, take, set, 0);
    }
    chunks[1].emit_op(Op::RETURN, 0);
    primitives::bundle::finalize_with_runtime_helpers(&mut chunks);
    primitives::globals::declare_free_globals(&mut chunks);
    primitives::globals::normalize_global_table(&mut chunks);
    chunks
}
fn worker(chunks: Vec<Chunk>) -> (VM, Value) {
    let mut vm = VM::new();
    primitives::platforms::register_platforms_all(&mut vm);
    let f = vm.run(chunks).unwrap();
    (vm, f)
}
fn args(single: bool, addr: i32, set: &[u8]) -> Vec<Value> {
    let mut a = vec![Value::I32(addr)];
    if single {
        a.push(Value::I32(set[0] as i32));
    }
    a
}
fn variants() -> Vec<(bool, Vec<u8>)> {
    vec![
        (true, b"a".to_vec()),
        (false, b"az".to_vec()),
        (false, b" \t\r\n".to_vec()),
        (false, b"01234567".to_vec()),
        (false, vec![1, 16, 31, 128, 143, 160, 254, 255]),
    ]
}
#[test]
fn masks_preserve_membership_nul_page_traps_and_inputs() {
    let mut checks = 0;
    let mut sizes = vec![];
    for (single, set) in variants() {
        for take in [false, true] {
            let programs = [
                program(false, take, &set, single),
                program(true, take, &set, single),
            ];
            for chunks in &programs {
                wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all())
                    .validate_all(&vybe_platform_wasm::write_wasm(chunks))
                    .unwrap();
            }
            let body = |chunks: &[Chunk]| chunks.iter().map(|c| c.code.len()).sum::<usize>();
            let native = |chunks: &[Chunk]| {
                vybe_platform_wasm::write_wasm_with_options(
                    chunks,
                    vybe_platform_wasm::WasmWriteOptions {
                        include_name_section: false,
                        include_vybe_metadata: false,
                    },
                )
                .len()
            };
            if single && !take {
                assert_eq!(body(&programs[0]), body(&programs[1]));
            } else {
                assert!(body(&programs[1]) < body(&programs[0]));
                assert!(native(&programs[1]) < native(&programs[0]));
            }
            let tag = set.iter().map(|b| format!("{b:02x}")).collect::<String>();
            if std::env::var_os("VYBE_PRIMITIVE_SPAN_EXPORT").is_some() {
                for side in 0..2 {
                    native_export::export_chunks(
                        &programs[side],
                        &format!("span-{tag}-{take}-{side}"),
                    );
                }
            }
            sizes.push(serde_json::json!({"set":set,"single":single,"take":take,"internal":[body(&programs[0]),body(&programs[1])],"native":[native(&programs[0]),native(&programs[1])],"locals":[programs[0].iter().map(|c|c.local_count as usize).sum::<usize>(),programs[1].iter().map(|c|c.local_count as usize).sum::<usize>()]}));
            let mut workers = programs.map(worker);
            for addr in [0usize, 1, 7, 15, 65503, 65520, 65535] {
                for length in [0usize, 1, 2, 7, 15, 16, 17, 31, 32, 63] {
                    for stop in [0u8, set[0], 0x91, 255] {
                        let results = workers.each_mut().map(|(vm, f)| {
                            vm.memory.with_buffer_mut(|m| {
                                m.fill(if take { set[0] } else { 0x91 });
                                if addr + length < m.len() {
                                    m[addr + length] = stop;
                                }
                                if addr + length + 1 < m.len() {
                                    m[addr + length + 1] = 0;
                                }
                            });
                            let before = vm.memory.with_buffer_mut(|m| m.to_vec());
                            let result = vm
                                .invoke(f, &args(single, addr as i32, &set))
                                .map(|v| v.as_f64())
                                .map_err(|e| e.message);
                            vm.memory.with_buffer_mut(|m| assert_eq!(m, &before[..]));
                            result
                        });
                        assert_eq!(
                            results[0], results[1],
                            "single={single} take={take} set={set:?} addr={addr} length={length} stop={stop}"
                        );
                        checks += 1;
                    }
                }
            }
            for value in 0..=255u8 {
                for lane in 0..16 {
                    for (vm, _) in &mut workers {
                        vm.memory.with_buffer_mut(|m| {
                            m.fill(if take { set[0] } else { 0x91 });
                            m[2 + lane] = value;
                            m[64] = 0;
                        });
                    }
                    let out = workers
                        .each_mut()
                        .map(|(vm, f)| vm.invoke(f, &args(single, 0, &set)).unwrap().as_f64());
                    assert_eq!(out[0], out[1]);
                    checks += 1;
                }
            }
        }
    }
    eprintln!("Verified {checks} span cases; sizes={sizes:?}");
    if std::env::var_os("VYBE_PRIMITIVE_SPAN_EXPORT").is_some() {
        std::fs::write(
            std::path::Path::new(env!("CARGO_MANIFEST_DIR")).join("../../SPAN_GATES.json"),
            serde_json::to_vec_pretty(&serde_json::json!({"checks":checks,"sizes":sizes})).unwrap(),
        )
        .unwrap();
    }
}
#[test]
#[ignore]
fn debug_pairs() {
    use std::time::{Duration, Instant};
    let mut rows = vec![];
    for (single, set) in variants() {
        for take in [false, true] {
            if single && !take {
                continue;
            }
            for length in [0usize, 8, 4096] {
                let mut workers = [
                    worker(program(false, take, &set, single)),
                    worker(program(true, take, &set, single)),
                ];
                for (vm, _) in &mut workers {
                    vm.memory.with_buffer_mut(|m| {
                        m.fill(if take { set[0] } else { 0x91 });
                        m[length] = 0;
                    });
                }
                let a = args(single, 0, &set);
                let mut samples = [vec![], vec![]];
                for pair in 0..7 {
                    for side in if pair % 2 == 0 { [0, 1] } else { [1, 0] } {
                        let (vm, f) = &mut workers[side];
                        let start = Instant::now();
                        let mut n = 0;
                        while start.elapsed() < Duration::from_millis(100) {
                            assert_eq!(
                                std::hint::black_box(vm.invoke(f, &a).unwrap()).as_f64(),
                                length as f64
                            );
                            n += 1;
                        }
                        samples[side].push(start.elapsed().as_secs_f64() / n as f64);
                    }
                }
                let mut ratios: Vec<_> = samples[0]
                    .iter()
                    .zip(&samples[1])
                    .map(|(a, b)| a / b)
                    .collect();
                ratios.sort_by(f64::total_cmp);
                rows.push(serde_json::json!({"single":single,"set":set,"take":take,"length":length,"median_speedup":ratios[3],"seconds_per_call":samples}));
            }
        }
    }
    let result = serde_json::json!({"profile":"opt0 with assertions and overflow checks","pairs":7,"minimum_sample_ms":100,"cases":rows});
    std::fs::write(
        std::path::Path::new(env!("CARGO_MANIFEST_DIR")).join("../../SPAN_MEASUREMENTS.json"),
        serde_json::to_vec_pretty(&result).unwrap(),
    )
    .unwrap();
    eprintln!("{result}");
}
