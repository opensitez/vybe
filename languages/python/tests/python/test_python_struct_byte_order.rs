use super::helpers::run_python;

// struct — byte order prefixes, format characters, calcsize, pack_into, unpack_from, iter_unpack, Struct class

#[test]
fn test_struct_little_endian_int32() {
    let out = run_python(r#"
import struct
data = struct.pack("<i", 1000)
print(len(data))
print(struct.unpack("<i", data)[0])
"#);
    assert_eq!(out, vec!["4", "1000"]);
}

#[test]
fn test_struct_big_endian_int32() {
    let out = run_python(r#"
import struct
data = struct.pack(">i", 1000)
print(len(data))
print(struct.unpack(">i", data)[0])
"#);
    assert_eq!(out, vec!["4", "1000"]);
}

#[test]
fn test_struct_little_vs_big_endian_bytes_differ() {
    let out = run_python(r#"
import struct
le = struct.pack("<i", 1)
be = struct.pack(">i", 1)
print(le != be)
print(le)
print(be)
"#);
    assert_eq!(out, vec!["True", "b'\\x01\\x00\\x00\\x00'", "b'\\x00\\x00\\x00\\x01'"]);
}

#[test]
fn test_struct_network_byte_order_exclamation() {
    let out = run_python(r#"
import struct
# '!' is network (big-endian)
data = struct.pack("!H", 80)
print(struct.unpack("!H", data)[0])
"#);
    assert_eq!(out, vec!["80"]);
}

#[test]
fn test_struct_padding_x_byte() {
    let out = run_python(r#"
import struct
# '2x' = 2 padding bytes, 'i' = int
data = struct.pack("2xi", 42)
print(len(data))
print(struct.unpack("2xi", data)[0])
"#);
    assert_eq!(out, vec!["6", "42"]);
}

#[test]
fn test_struct_calcsize_format() {
    let out = run_python(r#"
import struct
print(struct.calcsize(">HHi"))   # 2+2+4
print(struct.calcsize(">B"))     # 1
print(struct.calcsize(">d"))     # 8
"#);
    assert_eq!(out, vec!["8", "1", "8"]);
}

#[test]
fn test_struct_format_unsigned_vs_signed() {
    let out = run_python(r#"
import struct
data = struct.pack(">b", -1)
print(struct.unpack(">b", data)[0])
print(struct.unpack(">B", data)[0])
"#);
    assert_eq!(out, vec!["-1", "255"]);
}

#[test]
fn test_struct_half_float_e_format() {
    let out = run_python(r#"
import struct
data = struct.pack(">e", 1.0)
print(len(data))
val = struct.unpack(">e", data)[0]
print(abs(val - 1.0) < 0.001)
"#);
    assert_eq!(out, vec!["2", "True"]);
}

#[test]
fn test_struct_pack_into_buffer() {
    let out = run_python(r#"
import struct
buf = bytearray(8)
struct.pack_into(">i", buf, 0, 100)
struct.pack_into(">i", buf, 4, 200)
print(struct.unpack_from(">i", buf, 0)[0])
print(struct.unpack_from(">i", buf, 4)[0])
"#);
    assert_eq!(out, vec!["100", "200"]);
}

#[test]
fn test_struct_unpack_from_with_offset() {
    let out = run_python(r#"
import struct
data = b"\x00\x00" + struct.pack(">H", 1234)
val = struct.unpack_from(">H", data, 2)[0]
print(val)
"#);
    assert_eq!(out, vec!["1234"]);
}

#[test]
fn test_struct_iter_unpack() {
    let out = run_python(r#"
import struct
data = struct.pack(">HHH", 10, 20, 30)
values = [v for (v,) in struct.iter_unpack(">H", data)]
print(values)
"#);
    assert_eq!(out, vec!["[10, 20, 30]"]);
}

#[test]
fn test_struct_class_pre_compiled() {
    let out = run_python(r#"
import struct
s = struct.Struct(">ii")
data = s.pack(1, 2)
print(s.unpack(data))
"#);
    assert_eq!(out, vec!["(1, 2)"]);
}

#[test]
fn test_struct_class_size_attribute() {
    let out = run_python(r#"
import struct
s = struct.Struct(">ii")
print(s.size)
"#);
    assert_eq!(out, vec!["8"]);
}

#[test]
fn test_struct_format_double() {
    let out = run_python(r#"
import struct
val = 3.141592653589793
data = struct.pack(">d", val)
print(abs(struct.unpack(">d", data)[0] - val) < 1e-15)
"#);
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_struct_format_string_bytes() {
    let out = run_python(r#"
import struct
data = struct.pack(">5s", b"hello")
print(data)
print(struct.unpack(">5s", data)[0])
"#);
    assert_eq!(out, vec!["b'hello'", "b'hello'"]);
}

#[test]
fn test_struct_format_pascal_string() {
    let out = run_python(r#"
import struct
# 'p' = Pascal string: first byte is length, rest is data
data = struct.pack(">5p", b"hi")
print(len(data))
print(struct.unpack(">5p", data)[0])
"#);
    assert_eq!(out, vec!["5", "b'hi'"]);
}

#[test]
fn test_struct_error_on_too_small_buffer() {
    let out = run_python(r#"
import struct
try:
    struct.unpack(">i", b"\x00\x00")
except struct.error:
    print("struct.error")
"#);
    assert_eq!(out, vec!["struct.error"]);
}

#[test]
fn test_struct_native_byte_order_at_sign() {
    let out = run_python(r#"
import struct
# '@' = native byte order with alignment
data = struct.pack("@i", 42)
val = struct.unpack("@i", data)[0]
print(val)
"#);
    assert_eq!(out, vec!["42"]);
}

#[test]
fn test_struct_format_bool() {
    let out = run_python(r#"
import struct
data = struct.pack(">??", True, False)
print(struct.unpack(">??", data))
"#);
    assert_eq!(out, vec!["(True, False)"]);
}

#[test]
fn test_struct_multiple_values_roundtrip() {
    let out = run_python(r#"
import struct
values = (255, -1, 3.14, b"xy")
fmt = ">Bi f 2s"
packed = struct.pack(fmt, *values)
result = struct.unpack(fmt, packed)
print(result[0])
print(result[1])
print(abs(result[2] - 3.14) < 0.001)
print(result[3])
"#);
    assert_eq!(out, vec!["255", "-1", "True", "b'xy'"]);
}
