use crate::helpers::{run_print, run_python_one};

#[test]
fn bytes_from_list_of_ascii() {
    assert_eq!(run_print("bytes([65, 66, 67])"), "b'ABC'");
}

#[test]
fn bytes_decode_ascii() {
    assert_eq!(run_print("b'hello'.decode('ascii')"), "hello");
}

#[test]
fn bytes_decode_utf8_multibyte() {
    assert_eq!(run_print("b'\\xc3\\xa9'.decode('utf-8')"), "é");
}

#[test]
fn str_encode_utf8() {
    assert_eq!(run_print("'é'.encode('utf-8')"), "b'\\xc3\\xa9'");
}

#[test]
fn bytes_hex_method() {
    assert_eq!(run_print("b'\\x00\\xff'.hex()"), "00ff");
}

#[test]
fn bytes_fromhex_roundtrip() {
    assert_eq!(run_print("bytes.fromhex('4869')"), "b'Hi'");
}

#[test]
fn bytes_length() {
    assert_eq!(run_print("len(b'abc')"), "3");
}

#[test]
fn bytes_indexing_returns_int() {
    assert_eq!(run_print("b'ABC'[0]"), "65");
}

#[test]
fn bytes_slice_returns_bytes() {
    assert_eq!(run_print("b'abcdef'[1:4]"), "b'bcd'");
}

#[test]
fn bytes_add_concatenation() {
    assert_eq!(run_print("b'a' + b'b'"), "b'ab'");
}

#[test]
fn bytes_mul_repeat() {
    assert_eq!(run_print("b'x' * 3"), "b'xxx'");
}

#[test]
fn bytes_startswith_prefix() {
    assert_eq!(run_print("b'hello'.startswith(b'he')"), "True");
}

#[test]
fn bytes_endswith_suffix() {
    assert_eq!(run_print("b'hello'.endswith(b'lo')"), "True");
}

#[test]
fn bytes_replace_subsequence() {
    assert_eq!(run_print("b'a-b-a'.replace(b'a', b'x')"), "b'x-b-x'");
}

#[test]
fn bytes_split_on_delimiter() {
    assert_eq!(run_print("b'a,b,c'.split(b',')"), "[b'a', b'b', b'c']");
}

#[test]
fn bytes_join_iterable() {
    assert_eq!(run_print("b','.join([b'a', b'b'])"), "b'a,b'");
}

#[test]
fn bytes_strip_whitespace() {
    assert_eq!(run_print("b'  hi  '.strip()"), "b'hi'");
}

#[test]
fn bytes_upper_ascii() {
    assert_eq!(run_print("b'abc'.upper()"), "b'ABC'");
}

#[test]
fn bytes_lower_ascii() {
    assert_eq!(run_print("b'ABC'.lower()"), "b'abc'");
}

#[test]
fn bytes_literal_with_escape() {
    assert_eq!(run_print("b'line\\n'"), "b'line\\n'");
}

#[test]
fn bytes_equality() {
    assert_eq!(run_print("b'x' == b'x'"), "True");
}

#[test]
fn bytes_inequality() {
    assert_eq!(run_print("b'x' != b'y'"), "True");
}

#[test]
fn bytes_in_operator() {
    assert_eq!(run_print("b'a' in b'cab'"), "True");
}

#[test]
fn bytes_not_in_operator() {
    assert_eq!(run_print("b'z' not in b'abc'"), "True");
}

#[test]
fn bytes_compare_lexicographic() {
    assert_eq!(run_print("b'a' < b'b'"), "True");
}

#[test]
fn bytes_find_sub() {
    assert_eq!(run_print("b'abcabc'.find(b'ca')"), "2");
}

#[test]
fn bytes_rfind_sub() {
    assert_eq!(run_print("b'abcabc'.rfind(b'ca')"), "5");
}

#[test]
fn bytes_count_sub() {
    assert_eq!(run_print("b'aaa'.count(b'a')"), "3");
}

#[test]
fn bytes_partition() {
    assert_eq!(run_print("b'a.b.c'.partition(b'.')"), "(b'a', b'.', b'b.c')");
}

#[test]
fn bytes_rpartition() {
    assert_eq!(run_print("b'a.b.c'.rpartition(b'.')"), "(b'a.b', b'.', b'c')");
}

#[test]
fn bytes_startswith_tuple_options() {
    assert_eq!(run_print("b'hello'.startswith((b'he', b'xy'))"), "True");
}

#[test]
fn bytes_decode_errors_replace() {
    assert_eq!(
        run_print("b'\\xff'.decode('ascii', errors='replace')"),
        ""
    );
}

#[test]
fn str_encode_ascii() {
    assert_eq!(run_print("'A'.encode('ascii')"), "b'A'");
}

#[test]
fn bytes_from_literal_equals_constructor() {
    assert_eq!(run_print("bytes(b'abc')"), "b'abc'");
}

#[test]
fn bytes_iterable_in_for_loop_sum() {
    assert_eq!(
        run_python_one("total = 0\nfor b in b'\\x01\\x02\\x03':\n total += b\nprint(total)\n"),
        "6"
    );
}

#[test]
fn bytes_list_comprehension_ord_values() {
    assert_eq!(
        run_print("[b for b in b'abc']"),
        "[97, 98, 99]"
    );
}

#[test]
fn bytes_mutable_bytearray_distinction() {
    assert_eq!(run_print("type(bytearray(b'a')).__name__"), "bytearray");
}

#[test]
fn bytearray_decode_same_as_bytes() {
    assert_eq!(run_print("bytearray(b'hi').decode()"), "hi");
}

#[test]
fn bytearray_append_byte() {
    assert_eq!(
        run_python_one("ba = bytearray(b'a')\nba.append(ord('b'))\nprint(bytes(ba))\n"),
        "b'ab'"
    );
}

#[test]
fn bytes_repr_format() {
    assert_eq!(run_print("repr(b'\\n')"), "b'\\n'");
}

#[test]
fn bytes_empty() {
    assert_eq!(run_print("len(b'')"), "0");
}

#[test]
fn bytes_bool_empty_false() {
    assert_eq!(run_print("bool(b'')"), "False");
}

#[test]
fn bytes_bool_nonempty_true() {
    assert_eq!(run_print("bool(b'\\x00')"), "True");
}

#[test]
fn bytes_hex_empty() {
    assert_eq!(run_print("b''.hex()"), "");
}

#[test]
fn bytes_fromhex_empty() {
    assert_eq!(run_print("bytes.fromhex('')"), "b''");
}
