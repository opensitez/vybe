use super::helpers::*;

macro_rules! runtime_case {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(run_python_one($src), $expected);
        }
    };
}

macro_rules! compile_case {
    ($name:ident, $src:expr) => {
        #[test]
        fn $name() {
            compile_ok($src);
        }
    };
}

compile_case!(bytearray_basic_compile, "b = bytearray(b'abc')\n");
compile_case!(bytearray_from_list_compile, "b = bytearray([65, 66, 67])\n");
compile_case!(bytearray_append_compile, "b = bytearray()\nb.append(65)\n");
compile_case!(bytearray_extend_compile, "b = bytearray(b'a')\nb.extend(b'bc')\n");
compile_case!(bytearray_pop_compile, "b = bytearray(b'abc')\nx = b.pop()\n");
compile_case!(bytearray_remove_compile, "b = bytearray([1, 2, 3])\nb.remove(2)\n");
compile_case!(bytearray_reverse_compile, "b = bytearray(b'abc')\nb.reverse()\n");
compile_case!(bytearray_clear_compile, "b = bytearray(b'abc')\nb.clear()\n");
compile_case!(bytearray_slice_compile, "b = bytearray(b'abcdef')\nx = b[1:4]\n");
compile_case!(bytearray_slice_assign_compile, "b = bytearray(b'abc')\nb[1:2] = b'XYZ'\n");
compile_case!(bytes_from_list_compile, "b = bytes([65, 66, 67])\n");
compile_case!(bytes_join_compile, "b = b'-'.join([b'a', b'b'])\n");
compile_case!(bytes_split_compile, "parts = b'a,b,c'.split(b',')\n");
compile_case!(bytes_replace_compile, "b = b'abc'.replace(b'b', b'X')\n");
compile_case!(bytes_find_compile, "i = b'banana'.find(b'na')\n");
compile_case!(bytes_index_compile, "i = b'banana'.index(b'na')\n");
compile_case!(bytes_startswith_compile, "ok = b'abc'.startswith(b'a')\n");
compile_case!(bytes_endswith_compile, "ok = b'abc'.endswith(b'c')\n");
compile_case!(bytes_decode_compile, "s = b'hello'.decode('utf-8')\n");
compile_case!(bytes_format_percent_compile, "b = b'%s' % b'abc'\n");
runtime_case!(bytes_len_runtime, "print(len(b'hello'))\n", "5");
runtime_case!(bytes_index_runtime, "print(b'abc'[1])\n", "98");
runtime_case!(bytes_slice_len_runtime, "print(len(b'abcdef'[1:4]))\n", "3");
runtime_case!(bytes_membership_runtime, "print(98 in b'abc')\n", "true");
compile_case!(memoryview_basic_compile, "m = memoryview(b'abc')\n");
compile_case!(memoryview_slice_compile, "m = memoryview(b'abcdef')\ns = m[1:4]\n");
compile_case!(memoryview_tobytes_compile, "m = memoryview(b'abc')\nb = m.tobytes()\n");
compile_case!(bytes_hex_compile, "h = b'abc'.hex()\n");
compile_case!(bytes_fromhex_compile, "b = bytes.fromhex('414243')\n");
compile_case!(bytearray_fromhex_compile, "b = bytearray.fromhex('414243')\n");