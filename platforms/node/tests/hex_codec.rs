use vybe_runtime::{Chunk, Op, VM};
fn roundtrip(input: &str) -> String {
    let mut c = Chunk::new("hex_codec");
    c.emit_string_const(input, 0);
    c.emit_string_const("hex", 0);
    let decode = c.add_import("node:buffer", "from");
    c.emit_call(decode, 2, 0);
    c.emit_string_const("hex", 0);
    let encode = c.add_import("node:buffer", "toString");
    c.emit_call(encode, 2, 0);
    c.emit_op(Op::RETURN, 0);
    let mut vm = VM::new();
    vybe_platform_node::buffer::register(&mut vm);
    vm.run(vec![c]).unwrap().to_string()
}
#[test]
fn hex_stops_at_invalid_pairs_without_slicing_utf8() {
    for (input, expected) in [
        ("", ""),
        ("aB00ff", "ab00ff"),
        ("abc", "ab"),
        ("abggcd", "ab"),
        ("ab0xcd", "ab"),
        (" ab", ""),
        ("ab cd", "ab"),
        ("é0", ""),
        ("0é", ""),
        ("ab🙂cd", "ab"),
        ("ab\0cd", "ab"),
    ] {
        assert_eq!(roundtrip(input), expected, "input={input:?}");
    }
}
#[test]
fn all_byte_values_encode_as_lowercase_pairs() {
    let input = (0..=255)
        .map(|byte| format!("{byte:02x}"))
        .collect::<String>();
    assert_eq!(roundtrip(&input), input);
}

#[test]
fn latin1_binary_and_ascii_encode_utf16_low_bytes() {
    for encoding in ["latin1", "binary", "ascii"] {
        for (input, expected) in [
            ("Ā", "00"),
            ("🙂", "3d42"),
            ("A🙂B", "413d4242"),
            ("東京", "71ac"),
        ] {
            let mut c = Chunk::new("latin1_codec");
            c.emit_string_const(input, 0);
            c.emit_string_const(encoding, 0);
            let from = c.add_import("node:buffer", "from");
            c.emit_call(from, 2, 0);
            c.emit_string_const("hex", 0);
            let encode = c.add_import("node:buffer", "toString");
            c.emit_call(encode, 2, 0);
            c.emit_op(Op::RETURN, 0);
            let mut vm = VM::new();
            vybe_platform_node::buffer::register(&mut vm);
            assert_eq!(
                vm.run(vec![c]).unwrap().to_string(),
                expected,
                "{encoding} {input:?}"
            );
        }
    }
}
