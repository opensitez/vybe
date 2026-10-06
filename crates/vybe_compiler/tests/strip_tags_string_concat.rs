use vybe_compiler::primitives::{bundle, globals, platforms};
use vybe_runtime::{Chunk, Op, VM, Value};
#[path = "fixtures/strip_tags_string_before.rs"]
mod before;

fn runner(old: bool, argc: u8) -> (VM, Value) {
    let mut helper = Chunk::new("<strip-tags>");
    helper.arity = argc;
    helper.local_count = argc as u16;
    for slot in 0..argc {
        helper.emit_op_u16(Op::LOCAL_GET, slot as u16, 0);
    }
    let mut body = vec![helper];
    if old {
        before::emit_strip_tags(&mut body, 0, argc, 0);
    } else {
        vybe_language_php::emitter::string_adapter::emit_strip_tags(&mut body, 0, argc, 0);
    }
    body[0].emit_op(Op::RETURN, 0);
    eprintln!(
        "strip_tags old={old} argc={argc}: bytes={} locals={}",
        body[0].code.len(),
        body[0].local_count
    );
    let mut root = Chunk::new("<strip-tags-root>");
    root.emit_op_u16(Op::REF_FUNC, 1, 0);
    root.emit(0, 0);
    root.emit_op(Op::RETURN, 0);
    let mut chunks = vec![root, body.remove(0)];
    bundle::finalize_with_runtime_helpers(&mut chunks);
    globals::declare_free_globals(&mut chunks);
    globals::normalize_global_table(&mut chunks);
    let native = wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all())
        .validate_all(&vybe_platform_wasm::write_wasm(&chunks));
    if old {
        // The captured baseline's alphanumeric-name branch has no value in
        // either arm but declares an externref result. Keep its full native
        // failure visible while comparing the original VM behavior.
        let error = native
            .err()
            .expect("frozen statement block must expose its native type error");
        assert!(
            error
                .to_string()
                .contains("expected externref but nothing on stack")
        );
        eprintln!("known frozen native statement-block error: {error}");
    } else {
        native.unwrap();
    }
    let mut vm = VM::new();
    platforms::register_platforms_all(&mut vm);
    let callable = vm.run(chunks).unwrap();
    (vm, callable)
}

#[test]
fn frozen_parity_for_tags_unicode_malformed_and_scalar_coercion() {
    for argc in [1, 2] {
        let mut old = runner(true, argc);
        let mut new = runner(false, argc);
        for subject in [
            "",
            "plain",
            "<b>bold</b>",
            "<B class='x'>a</B><i>b</i>",
            "<!-- comment -->x<!doctype html>y",
            "a<unfinished",
            "1 < 2 > 0",
            "😀<b>é\u{301}</b>🦀",
            "a\0<b>b</b>",
            "<<b>nested</b>>",
        ] {
            for allowed in ["", "<b>", "<B><i>"] {
                let mut args = vec![Value::String(subject.into())];
                if argc == 2 {
                    args.push(Value::String(allowed.into()));
                }
                let expected = old.0.invoke(&old.1, &args).unwrap();
                let actual = new.0.invoke(&new.1, &args).unwrap();
                assert!(matches!(actual, Value::String(_)));
                assert_eq!(actual, expected, "subject={subject:?} allowed={allowed:?}");
            }
        }
        for subject in [
            Value::Null,
            Value::Bool(false),
            Value::Bool(true),
            Value::I32(42),
            Value::F64(1.5),
        ] {
            let mut args = vec![subject];
            if argc == 2 {
                args.push(Value::Null);
            }
            assert_eq!(
                new.0.invoke(&new.1, &args).unwrap(),
                old.0.invoke(&old.1, &args).unwrap()
            );
        }
    }
}

#[test]
fn paired_debug_strip_tags_measurements() {
    for subject in ["plain text ".repeat(96), "<b>text</b><i>x</i>".repeat(48)] {
        let mut runners = [runner(true, 2), runner(false, 2)];
        let args = [Value::String(subject.into()), Value::String("<b>".into())];
        let old_callable = runners[0].1.clone();
        let expected = runners[0].0.invoke(&old_callable, &args).unwrap();
        let measure = |runner: &mut (VM, Value), n: usize| {
            let t = std::time::Instant::now();
            let mut last = Value::Undefined;
            for _ in 0..n {
                last = std::hint::black_box(runner.0.invoke(&runner.1, &args).unwrap());
            }
            let elapsed = t.elapsed().as_secs_f64();
            assert_eq!(last, expected);
            elapsed
        };
        let mut n = 1;
        while measure(&mut runners[1], n) < 0.03 {
            n *= 2;
        }
        let mut ratios = Vec::new();
        for round in 0..7 {
            let mut times = [0.0; 2];
            for index in [round % 2, 1 - round % 2] {
                times[index] = measure(&mut runners[index], n);
            }
            ratios.push(times[0] / times[1]);
            eprintln!(
                "strip_tags debug pair {round}: n={n} before={}ms after={}ms ratio={}",
                times[0] * 1000.0,
                times[1] * 1000.0,
                ratios.last().unwrap()
            );
        }
        ratios.sort_by(f64::total_cmp);
        eprintln!(
            "strip_tags paired median={}x range={}..{}",
            ratios[3], ratios[0], ratios[6]
        );
    }
}
