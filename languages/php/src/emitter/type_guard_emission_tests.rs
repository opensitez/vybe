use super::*;

#[test]
fn guards_spill_only_the_suffix_and_keep_the_checked_argument() {
    for argc in 1..=5u8 {
        for arg_idx in 0..argc {
            for expect in [Expect::Array, Expect::NotArray, Expect::NotScalar, Expect::StringCoercible] {
                let base = 11u16;
                let mut chunks = vec![Chunk::new("guard".to_string())];
                chunks[0].local_count = base;
                guard_arg(&mut chunks, 0, argc, arg_idx, expect, "TypeError", "bad argument", 7);

                let mut prefix = Chunk::new("prefix".to_string());
                for i in (arg_idx..argc).rev() {
                    prefix.emit_op_u16(
                        if i == arg_idx { Op::LOCAL_TEE } else { Op::LOCAL_SET },
                        base + u16::from(i - arg_idx),
                        7,
                    );
                }
                assert!(chunks[0].code.starts_with(&prefix.code));

                let mut suffix = Chunk::new("suffix".to_string());
                for i in 1..(argc - arg_idx) {
                    suffix.emit_op_u16(Op::LOCAL_GET, base + u16::from(i), 7);
                }
                assert!(chunks[0].code.ends_with(&suffix.code));
                // The existing cold throw path retains its one message slot.
                assert_eq!(chunks[0].local_count, base + u16::from(argc - arg_idx) + 1);
            }
        }
    }
}

#[test]
fn top_argument_guard_uses_one_setup_instruction_and_no_reload() {
    let mut chunks = vec![Chunk::new("guard".to_string())];
    guard_arg(&mut chunks, 0, 5, 4, Expect::Array, "TypeError", "bad argument", 13);
    let mut prefix = Chunk::new("prefix".to_string());
    prefix.emit_op_u16(Op::LOCAL_TEE, 0, 13);
    assert!(chunks[0].code.starts_with(&prefix.code));
    assert_eq!(chunks[0].local_count, 2);
}
