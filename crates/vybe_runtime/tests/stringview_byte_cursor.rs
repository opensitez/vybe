use std::sync::Arc;
use vybe_runtime::{
    Chunk, Op, VM, Value,
    value::{Object, ObjectKind},
};

fn view(text: &str, legacy: bool) -> Value {
    let mut create = Chunk::new("create-string-iterator");
    create.emit_string_const(text, 0);
    create.emit_op(Op::STRING_AS_ITER, 0);
    create.emit_op(Op::RETURN, 0);
    let value = VM::new().run(vec![create]).unwrap();
    if legacy {
        let Value::Object(object) = &value else {
            unreachable!()
        };
        object.lock().unwrap().kind = ObjectKind::Ordinary;
    }
    value
}

fn worker(op: Op) -> (VM, Value) {
    let mut root = Chunk::new("iterator-worker-owner");
    root.emit_op_u16(Op::REF_FUNC, 1, 0);
    root.emit(0, 0);
    root.emit_op(Op::RETURN, 0);
    let mut code = Chunk::new("iterator-operation");
    code.arity = 2;
    code.param_count = 2;
    code.local_count = 2;
    code.emit_op_u16(Op::LOCAL_GET, 0, 0);
    if op != Op::STRINGVIEW_ITER_NEXT {
        code.emit_op_u16(Op::LOCAL_GET, 1, 0);
    }
    code.emit_op(op, 0);
    code.emit_op(Op::RETURN, 0);
    let mut vm = VM::new();
    let function = vm.run(vec![root, code]).unwrap();
    (vm, function)
}

fn position(view: &Value) -> Value {
    let Value::Object(object) = view else {
        unreachable!()
    };
    object
        .lock()
        .unwrap()
        .properties
        .get("__iter_pos")
        .cloned()
        .unwrap_or(Value::Undefined)
}

#[test]
fn mixed_unicode_moves_and_slices_match_scalar_reference_and_legacy_views() {
    let ops = [
        Op::STRINGVIEW_ITER_NEXT,
        Op::STRINGVIEW_ITER_ADVANCE,
        Op::STRINGVIEW_ITER_REWIND,
        Op::STRINGVIEW_ITER_SLICE,
    ];
    let mut workers = ops.map(worker);
    let mut texts: Vec<String> = ["", "a", "a\0é中😀𐐀z", "😀😀😀", "e\u{301}\r\n"]
        .into_iter()
        .map(str::to_owned)
        .collect();
    let mut seed = 0x62e2_46a9u32;
    for _ in 0..256 {
        let text = (0..32)
            .filter_map(|_| {
                seed = seed.wrapping_mul(1664525).wrapping_add(1013904223);
                char::from_u32(seed % 0x110000)
            })
            .collect();
        texts.push(text);
    }
    let mut checks = 0;
    for text in texts {
        let chars: Vec<char> = text.chars().collect();
        let views = [view(&text, false), view(&text, true)];
        let mut pos = 0;
        for step in 0..160 {
            seed = seed.wrapping_mul(1664525).wrapping_add(1013904223);
            let index = (seed >> 16) as usize % 4;
            let n = [0, 1, 2, 3, 17, u32::MAX][step % 6] as usize;
            let expected = match index {
                0 => {
                    let cp = chars.get(pos).map_or(-1, |c| *c as i32);
                    if pos < chars.len() {
                        pos += 1;
                    }
                    Value::I32(cp)
                }
                1 => {
                    let moved = n.min(chars.len() - pos);
                    pos += moved;
                    Value::I32(moved as i32)
                }
                2 => {
                    let moved = n.min(pos);
                    pos -= moved;
                    Value::I32(moved as i32)
                }
                3 => Value::String(chars[pos..].iter().take(n).collect::<String>().into()),
                _ => unreachable!(),
            };
            let (vm, function) = &mut workers[index];
            for view in &views {
                let actual = vm
                    .invoke(function, &[view.clone(), Value::I32(n as u32 as i32)])
                    .unwrap();
                assert_eq!(
                    actual, expected,
                    "text={text:?} step={step} op={index} n={n}"
                );
                assert_eq!(position(view), Value::I32(pos as i32));
                checks += 1;
            }
        }
    }
    eprintln!("cursor scalar/legacy comparisons={checks}");
}

#[test]
fn metadata_edits_fall_back_and_aliases_share_the_live_cursor() {
    let mut next = worker(Op::STRINGVIEW_ITER_NEXT);
    let candidate = view("aé😀z", false);
    let alias = candidate.clone();
    for (object, expected) in [(&candidate, 'a'), (&alias, 'é')] {
        assert_eq!(
            next.0
                .invoke(&next.1, &[object.clone(), Value::I32(0)])
                .unwrap(),
            Value::I32(expected as i32)
        );
    }
    let Value::Object(object) = &candidate else {
        unreachable!()
    };
    object
        .lock()
        .unwrap()
        .properties
        .insert("__iter_pos".into(), Value::I32(0));
    assert_eq!(
        next.0
            .invoke(&next.1, &[alias.clone(), Value::I32(0)])
            .unwrap(),
        Value::I32('a' as i32)
    );
    {
        let mut object = object.lock().unwrap();
        object
            .properties
            .insert("__iter_str".into(), Value::String(Arc::from("中Ω")));
        object.properties.insert("__iter_pos".into(), Value::I32(1));
    }
    assert_eq!(
        next.0
            .invoke(&next.1, &[alias.clone(), Value::I32(0)])
            .unwrap(),
        Value::I32('Ω' as i32)
    );
    assert_eq!(
        next.0.invoke(&next.1, &[candidate, Value::I32(0)]).unwrap(),
        Value::I32(-1)
    );
}

#[test]
fn null_and_invalid_view_traps_match_the_compatibility_path() {
    for op in [
        Op::STRINGVIEW_ITER_NEXT,
        Op::STRINGVIEW_ITER_ADVANCE,
        Op::STRINGVIEW_ITER_REWIND,
        Op::STRINGVIEW_ITER_SLICE,
    ] {
        let (mut vm, function) = worker(op);
        for value in [
            Value::Null,
            Value::Undefined,
            Value::Object(Arc::new(std::sync::Mutex::new(Object::new()))),
        ] {
            assert!(
                vm.invoke(&function, &[value, Value::I32(1)])
                    .unwrap_err()
                    .is_trap()
            );
        }
    }
}

fn walk_worker(count: usize) -> (VM, Value) {
    let mut root = Chunk::new("walk-owner");
    root.emit_op_u16(Op::REF_FUNC, 1, 0);
    root.emit(0, 0);
    root.emit_op(Op::RETURN, 0);
    let mut code = Chunk::new("walk");
    code.arity = 1;
    code.param_count = 1;
    code.local_count = 2;
    // Reset inside every invocation; repeated runs walk the same complete view.
    code.emit_op_u16(Op::LOCAL_GET, 0, 0);
    code.emit_i32_const(-1, 0);
    code.emit_op(Op::STRINGVIEW_ITER_REWIND, 0);
    code.emit_op(Op::DROP, 0);
    code.emit_i32_const(0, 0);
    code.emit_op_u16(Op::LOCAL_SET, 1, 0);
    for _ in 0..count {
        code.emit_op_u16(Op::LOCAL_GET, 1, 0);
        code.emit_op_u16(Op::LOCAL_GET, 0, 0);
        code.emit_op(Op::STRINGVIEW_ITER_NEXT, 0);
        code.emit_op(Op::I32_ADD, 0);
        code.emit_op_u16(Op::LOCAL_SET, 1, 0);
    }
    code.emit_op_u16(Op::LOCAL_GET, 1, 0);
    code.emit_op(Op::RETURN, 0);
    let mut vm = VM::new();
    let function = vm.run(vec![root, code]).unwrap();
    (vm, function)
}

#[test]
fn complete_walks_preserve_every_codepoint_including_nul_and_non_bmp() {
    for count in [32, 1024] {
        let text: String = "a\0é中😀\u{10ffff}".chars().cycle().take(count).collect();
        let expected = text.chars().fold(0i32, |sum, c| sum.wrapping_add(c as i32));
        let (mut vm, function) = walk_worker(count);
        for legacy in [false, true] {
            let view = view(&text, legacy);
            assert_eq!(
                vm.invoke(&function, &[view.clone()]).unwrap(),
                Value::I32(expected)
            );
            assert_eq!(position(&view), Value::I32(count as i32));
        }
    }
}

#[test]
#[ignore = "paired complete-walk debug runtime measurement"]
fn paired_complete_walks() {
    use std::time::Instant;
    for count in [32, 1024, 16384] {
        for pattern in ["a", "é", "中", "😀", "aé中😀"] {
            let text: String = pattern.chars().cycle().take(count).collect();
            let expected = text.chars().fold(0i32, |sum, c| sum.wrapping_add(c as i32));
            let views = [view(&text, true), view(&text, false)];
            let (mut vm, function) = walk_worker(count);
            for view in &views {
                assert_eq!(
                    vm.invoke(&function, &[view.clone()]).unwrap(),
                    Value::I32(expected)
                );
            }
            let mut pairs = Vec::new();
            for pair in 0..7 {
                let mut times = [0.0; 2];
                for index in if pair % 2 == 0 { [0, 1] } else { [1, 0] } {
                    let start = Instant::now();
                    let result = vm.invoke(&function, &[views[index].clone()]).unwrap();
                    times[index] = start.elapsed().as_secs_f64() * 1000.0;
                    assert_eq!(result, Value::I32(expected));
                }
                pairs.push(times);
            }
            eprintln!("CURSOR_WALK count={count} pattern={pattern:?} pairs_ms={pairs:?}");
        }
    }
}
