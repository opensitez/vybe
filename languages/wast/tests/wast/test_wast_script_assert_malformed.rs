use super::helpers::parse_ok;

#[test]
fn assert_malformed_unclosed_string() {
    parse_ok(r#"(assert_malformed (module quote "(module (func (export \"f\"))") "unexpected token")"#);
}

#[test]
fn assert_malformed_unclosed_block() {
    parse_ok(r#"(assert_malformed (module quote "(module (func (block ") "unexpected token")"#);
}

#[test]
fn assert_malformed_invalid_character() {
    parse_ok(r#"(assert_malformed (module quote "(module (func @))") "unexpected token")"#);
}

#[test]
fn assert_malformed_unknown_opcode() {
    parse_ok(r#"(assert_malformed (module quote "(module (func invalid.opcode))") "unknown operator")"#);
}

#[test]
fn assert_malformed_invalid_integer() {
    parse_ok(r#"(assert_malformed (module quote "(module (func i32.const 9999999999999999999999999))") "constant out of range")"#);
}

#[test]
fn assert_malformed_invalid_float() {
    parse_ok(r#"(assert_malformed (module quote "(module (func f32.const invalid))") "unknown operator")"#);
}

#[test]
fn assert_malformed_invalid_utf8() {
    parse_ok(r#"(assert_malformed (module quote "(module (data \"\\ff\"))") "invalid utf-8 encoding")"#);
}

#[test]
fn assert_malformed_binary_magic() {
    parse_ok(r#"(assert_malformed (module binary "\00asm") "magic header not detected")"#);
}

#[test]
fn assert_malformed_binary_version() {
    parse_ok(r#"(assert_malformed (module binary "\00asm\00\00\00\00") "unknown binary version")"#);
}

#[test]
fn assert_malformed_binary_invalid_section() {
    parse_ok(r#"(assert_malformed (module binary "\00asm\01\00\00\00\ff\00") "malformed section id")"#);
}

#[test]
fn assert_malformed_binary_unexpected_eof() {
    parse_ok(r#"(assert_malformed (module binary "\00asm\01\00\00\00\01\01") "unexpected end")"#);
}
