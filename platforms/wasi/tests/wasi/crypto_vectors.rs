use std::sync::Arc;

use vybe_bytecode::{Chunk, Op, VM, Value};
use vybe_bytecode::capabilities::Capabilities;
use vybe_compiler::compiler::platforms::register_platforms;

fn call_import(module: &str, name: &str, args: Vec<Value>) -> Value {
    let mut chunk = Chunk::new("<wasi-crypto-vectors-test>");
    let import_idx = chunk.add_import(module, name);
    let argc = args.len() as u8;
    for value in args {
        let constant = chunk.add_constant(value);
        chunk.emit_op_u16(Op::CONST, constant, 0);
    }
    chunk.emit_op_u16(Op::CALL_IMPORT, import_idx, 0);
    chunk.emit(argc, 0);
    chunk.emit_op(Op::RETURN, 0);

    let mut vm = VM::new();
    register_platforms(&mut vm, &Capabilities::all());
    vm.run(vec![chunk]).expect("VM run failed")
}

fn s(text: &str) -> Value {
    Value::String(Arc::from(text))
}

macro_rules! md5_vector_test {
    ($name:ident, $input:expr, $expected:expr) => {
        #[test]
        fn $name() {
            let digest = call_import("wasi:crypto/hashes", "md5", vec![s($input)]);
            assert_eq!(digest, s($expected));
        }
    };
}

macro_rules! sha256_vector_test {
    ($name:ident, $input:expr, $expected:expr) => {
        #[test]
        fn $name() {
            let digest = call_import("wasi:crypto/hashes", "sha256", vec![s($input)]);
            assert_eq!(digest, s($expected));
        }
    };
}

md5_vector_test!(
    md5_of_single_letter_a_matches_rfc_vector,
    "a",
    "0cc175b9c0f1b6a831c399e269772661"
);
md5_vector_test!(
    md5_of_message_digest_matches_rfc_vector,
    "message digest",
    "f96b697d7cb7938d525a2f31aaf161d0"
);
md5_vector_test!(
    md5_of_lowercase_alphabet_matches_rfc_vector,
    "abcdefghijklmnopqrstuvwxyz",
    "c3fcd3d76192e4007dfb496cca67e13b"
);
md5_vector_test!(
    md5_of_alnum_sequence_matches_rfc_vector,
    "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789",
    "d174ab98d277d9f5a5611c2c9f419d9f"
);
md5_vector_test!(
    md5_of_numeric_80_byte_vector_matches_rfc_vector,
    "12345678901234567890123456789012345678901234567890123456789012345678901234567890",
    "57edf4a22be3c955ac49da2e2107b67a"
);
md5_vector_test!(
    md5_of_quick_brown_fox_matches_common_vector,
    "The quick brown fox jumps over the lazy dog",
    "9e107d9d372bb6826bd81d3542a419d6"
);
md5_vector_test!(
    md5_of_quick_brown_fox_with_period_matches_common_vector,
    "The quick brown fox jumps over the lazy dog.",
    "e4d909c290d0fb1ca068ffaddf22cbd0"
);
md5_vector_test!(
    md5_of_multiblock_ascii_sequence_matches_known_vector,
    "abcdbcdecdefdefgefghfghighijhijkijkljklmklmnlmnomnopnopq",
    "8215ef0796a20bcaaae116d3876c664a"
);
md5_vector_test!(
    md5_of_component_model_identifier_matches_known_vector,
    "component-model",
    "849de7d604b3247458048bfd7b33af05"
);

sha256_vector_test!(
    sha256_of_single_letter_a_matches_known_vector,
    "a",
    "ca978112ca1bbdcafac231b39a23dc4da786eff8147c4e72b9807785afee48bb"
);
sha256_vector_test!(
    sha256_of_message_digest_matches_known_vector,
    "message digest",
    "f7846f55cf23e14eebeab5b4e1550cad5b509e3348fbc4efa3a1413d393cb650"
);
sha256_vector_test!(
    sha256_of_lowercase_alphabet_matches_known_vector,
    "abcdefghijklmnopqrstuvwxyz",
    "71c480df93d6ae2f1efad1447c66c9525e316218cf51fc8d9ed832f2daf18b73"
);
sha256_vector_test!(
    sha256_of_alnum_sequence_matches_known_vector,
    "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789",
    "db4bfcbd4da0cd85a60c3c37d3fbd8805c77f15fc6b1fdfe614ee0a7c8fdb4c0"
);
sha256_vector_test!(
    sha256_of_numeric_80_byte_vector_matches_known_vector,
    "12345678901234567890123456789012345678901234567890123456789012345678901234567890",
    "f371bc4a311f2b009eef952dd83ca80e2b60026c8e935592d0f9c308453c813e"
);
sha256_vector_test!(
    sha256_of_hello_world_matches_known_vector,
    "hello world",
    "b94d27b9934d3e08a52e52d7da7dabfac484efe37a5380ee9088f7ace2efcde9"
);
sha256_vector_test!(
    sha256_of_quick_brown_fox_matches_common_vector,
    "The quick brown fox jumps over the lazy dog",
    "d7a8fbb307d7809469ca9abcb0082e4f8d5651e46d3cdb762d02d0bf37c9e592"
);
sha256_vector_test!(
    sha256_of_quick_brown_fox_with_period_matches_common_vector,
    "The quick brown fox jumps over the lazy dog.",
    "ef537f25c895bfa782526529a9b63d97aa631564d5d789c2b765448c8635fb6c"
);
sha256_vector_test!(
    sha256_of_multiblock_ascii_sequence_matches_known_vector,
    "abcdbcdecdefdefgefghfghighijhijkijkljklmklmnlmnomnopnopq",
    "248d6a61d20638b8e5c026930c3e6039a33ce45964ff2167f6ecedd419db06c1"
);
sha256_vector_test!(
    sha256_of_component_model_identifier_matches_known_vector,
    "component-model",
    "d81b486adf2062d06f8322de7af83fa8963bad4d016213976eb852d5f36fb170"
);

md5_vector_test!(
    md5_of_numeric_input_matches_stringified_value,
    "42",
    "a1d0c6e83f027327d8461063f4ac58a6"
);
md5_vector_test!(
    md5_of_boolean_true_input_matches_stringified_value,
    "true",
    "b326b5062b2f0e69046810717534cb09"
);
md5_vector_test!(
    md5_of_null_input_matches_stringified_value,
    "null",
    "37a6259cc0c1dae299a7866489dff0bd"
);
md5_vector_test!(
    md5_of_decimal_input_matches_stringified_value,
    "3.14",
    "4beed3b9c4a886067de0e3a094246f78"
);
md5_vector_test!(
    md5_of_boolean_false_input_matches_stringified_value,
    "false",
    "68934a3e9455fa72420237eb05902327"
);
md5_vector_test!(
    md5_of_negative_number_input_matches_stringified_value,
    "-7",
    "74687a12d3915d3c4d83f1af7b3683d5"
);

sha256_vector_test!(
    sha256_of_numeric_input_matches_stringified_value,
    "42",
    "73475cb40a568e8da8a045ced110137e159f890ac4da883b6b17dc651b3a8049"
);
sha256_vector_test!(
    sha256_of_boolean_true_input_matches_stringified_value,
    "true",
    "b5bea41b6c623f7c09f1bf24dcae58ebab3c0cdd90ad966bc43a45b44867e12b"
);
sha256_vector_test!(
    sha256_of_null_input_matches_stringified_value,
    "null",
    "74234e98afe7498fb5daf1f36ac2d78acc339464f950703b8c019892f982b90b"
);
sha256_vector_test!(
    sha256_of_decimal_input_matches_stringified_value,
    "3.14",
    "2efff1261c25d94dd6698ea1047f5c0a7107ca98b0a6c2427ee6614143500215"
);
sha256_vector_test!(
    sha256_of_boolean_false_input_matches_stringified_value,
    "false",
    "fcbcf165908dd18a9e49f7ff27810176db8e9f63b4352213741664245224f8aa"
);
sha256_vector_test!(
    sha256_of_negative_number_input_matches_stringified_value,
    "-7",
    "a770d3270c9dcdedf12ed9fd70444f7c8a95c26cae3cae9bd867499090a2f14b"
);
