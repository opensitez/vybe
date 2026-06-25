use crate::helpers::{run_python_one, run_print};

#[test]
fn bytes_literal_ascii() {
    assert_eq!(run_print("b'hi'"), "b'hi'");
}

#[test]
fn bytes_len() {
    assert_eq!(run_print("len(b'abc')"), "3");
}

#[test]
fn bytes_index() {
    assert_eq!(run_print("b'abc'[0]"), "97");
}

#[test]
fn bytes_slice() {
    assert_eq!(run_print("b'hello'[1:4]"), "b'ell'");
}

#[test]
fn bytes_concat() {
    assert_eq!(run_print("b'a' + b'b'"), "b'ab'");
}

#[test]
fn bytes_repeat() {
    assert_eq!(run_print("b'x' * 3"), "b'xxx'");
}

#[test]
fn bytes_startswith() {
    assert_eq!(run_print("b'hello'.startswith(b'he')"), "True");
}

#[test]
fn bytes_endswith() {
    assert_eq!(run_print("b'hello'.endswith(b'lo')"), "True");
}

#[test]
fn bytes_decode_utf8() {
    assert_eq!(run_print("b'hi'.decode()"), "hi");
}

#[test]
fn bytes_encode_str() {
    assert_eq!(run_print("'hi'.encode()"), "b'hi'");
}

#[test]
fn bytes_hex_escape() {
    assert_eq!(run_print("b'\\xff'"), "b'\\xff'");
}

#[test]
fn bytes_from_list_of_ints() {
    assert_eq!(run_print("bytes([65, 66])"), "b'AB'");
}

#[test]
fn bytes_to_list() {
    assert_eq!(run_print("list(b'ab')"), "[97, 98]");
}

#[test]
fn bytes_in_membership() {
    assert_eq!(run_print("b'a' in b'abc'"), "True");
}

#[test]
fn bytes_equality() {
    assert_eq!(run_print("b'abc' == b'abc'"), "True");
}

#[test]
fn bytes_compare_ordering() {
    assert_eq!(run_print("b'a' < b'b'"), "True");
}

#[test]
fn bytes_split() {
    assert_eq!(run_print("b'a,b'.split(b',')"), "[b'a', b'b']");
}

#[test]
fn bytes_join() {
    assert_eq!(run_print("b','.join([b'a', b'b'])"), "b'a,b'");
}

#[test]
fn bytes_strip() {
    assert_eq!(run_print("b'  x  '.strip()"), "b'x'");
}

#[test]
fn bytes_replace() {
    assert_eq!(run_print("b'aba'.replace(b'a', b'x')"), "b'xbx'");
}

#[test]
fn bytes_find_sub() {
    assert_eq!(run_print("b'abcabc'.find(b'ca')"), "2");
}

#[test]
fn bytes_count_sub() {
    assert_eq!(run_print("b'aaa'.count(b'a')"), "3");
}

#[test]
fn bytes_upper_ascii() {
    assert_eq!(run_print("b'ab'.upper()"), "b'AB'");
}

#[test]
fn bytes_lower_ascii() {
    assert_eq!(run_print("b'AB'.lower()"), "b'ab'");
}

#[test]
fn bytes_isalpha_ascii() {
    assert_eq!(run_print("b'abc'.isalpha()"), "True");
}

#[test]
fn bytes_isdigit_ascii() {
    assert_eq!(run_print("b'123'.isdigit()"), "True");
}

#[test]
fn bytes_hex_method() {
    assert_eq!(run_print("b'\\xff'.hex()"), "ff");
}

#[test]
fn bytes_fromhex() {
    assert_eq!(run_print("bytes.fromhex('4869')"), "b'Hi'");
}

#[test]
fn bytearray_mutable_setitem() {
    assert_eq!(
        run_python_one("ba = bytearray(b'abc')\nba[0] = ord('x')\nprint(ba)\n"),
        "bytearray(b'xbc')"
    );
}

#[test]
fn bytearray_append() {
    assert_eq!(
        run_python_one("ba = bytearray(b'a')\nba.append(ord('b'))\nprint(ba)\n"),
        "bytearray(b'ab')"
    );
}

#[test]
fn bytearray_extend() {
    assert_eq!(
        run_python_one("ba = bytearray(b'a')\nba.extend(b'bc')\nprint(ba)\n"),
        "bytearray(b'abc')"
    );
}

#[test]
fn bytearray_decode() {
    assert_eq!(run_print("bytearray(b'hi').decode()"), "hi");
}

#[test]
fn memoryview_from_bytes() {
    assert_eq!(
        run_python_one("mv = memoryview(b'abc')\nprint(len(mv))\n"),
        "3"
    );
}

#[test]
fn bytes_iter_values() {
    assert_eq!(run_print("list(b'ab')"), "[97, 98]");
}

#[test]
fn bytes_partition() {
    assert_eq!(run_print("b'a-b-c'.partition(b'-')"), "(b'a', b'-', b'b-c')");
}

#[test]
fn bytes_rpartition() {
    assert_eq!(run_print("b'a-b-c'.rpartition(b'-')"), "(b'a-b', b'-', b'c')");
}

#[test]
fn bytes_startswith_tuple() {
    assert_eq!(run_print("b'hello'.startswith((b'he', b'xy'))"), "True");
}

#[test]
fn bytes_maketrans_translate() {
    assert_eq!(
        run_print("b'abc'.translate(bytes.maketrans(b'a', b'x'))"),
        "b'xbc'"
    );
}

#[test]
fn bytes_center() {
    assert_eq!(run_print("b'ab'.center(4, b'-')"), "b'-ab-'");
}

#[test]
fn bytes_zfill() {
    assert_eq!(run_print("b'42'.zfill(5)"), "b'00042'");
}

#[test]
fn bytes_removeprefix() {
    assert_eq!(run_print("b'pre_x'.removeprefix(b'pre_')"), "b'x'");
}

#[test]
fn bytes_removesuffix() {
    assert_eq!(run_print("b'x_suf'.removesuffix(b'_suf')"), "b'x'");
}

#[test]
fn bytes_splitlines() {
    assert_eq!(run_print("b'a\\nb'.splitlines()"), "[b'a', b'b']");
}

#[test]
fn bytes_rsplit_maxsplit() {
    assert_eq!(run_print("b'a b c'.rsplit(b' ', 1)"), "[b'a b', b'c']");
}

#[test]
fn bytes_capitalize_ascii() {
    assert_eq!(run_print("b'hello'.capitalize()"), "b'Hello'");
}

#[test]
fn bytes_title_ascii() {
    assert_eq!(run_print("b'hello world'.title()"), "b'Hello World'");
}
