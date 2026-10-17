use super::*;

#[test]
fn multi_result_branches_keep_values_references_and_label_cleanup() {
    let reference: Arc<str> = Arc::from("branch result identity");
    for arity in [3, 4, 16, 255] {
        for gap in [0, 1, 8] {
            for is_loop in [false, true] {
                for depth in 0..3 {
                    let mut script = Chunk::new("<branch-results>");
                    script.emit_i32_const(0, 0);
                    script.emit_op(Op::RETURN, 0);
                    let mut vm = VM::new();
                    let entry = LabelEntry {
                        target: 0, is_loop, is_try: false, try_group: 0,
                        result_arity: arity as u8, stack_height: 2,
                    };
                    vm.label_stack = vec![entry; 3];
                    vm.stack = vec![Value::I32(11), Value::I64(22)];
                    let mut expected = vm.stack.clone();
                    vm.stack.extend((0..gap).map(|i| Value::I32(-1 - i)));
                    let results: Vec<_> = (0..arity).map(|i| match i % 4 {
                        0 => Value::String(reference.clone()),
                        1 => Value::I64(i as i64),
                        2 => Value::V128([i as u8; 16]),
                        _ => Value::I32(i as i32),
                    }).collect();
                    expected.extend(results.iter().cloned());
                    vm.stack.extend(results);
                    vm.branch_to_label(depth, entry);
                    assert_eq!(vm.stack, expected, "arity {arity}, gap {gap}, depth {depth}, loop {is_loop}");
                    assert_eq!(vm.label_stack.len(), if is_loop { 3 - depth } else { 2 - depth });
                    for result in &vm.stack[2..] {
                        if let Value::String(value) = result { assert!(Arc::ptr_eq(value, &reference)); }
                    }
                }
            }
        }
    }
    assert_eq!(Arc::strong_count(&reference), 1, "no branch result reference is leaked");
}
