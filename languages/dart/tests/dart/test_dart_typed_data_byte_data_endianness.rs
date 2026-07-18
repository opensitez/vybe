use super::helpers::run_prints;

// ═══════════════════════════════════════════════════════════
// Dart: dart:typed_data ByteData & Endianness
// ═══════════════════════════════════════════════════════════

#[test]
fn byte_data_creation() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:typed_data';
void main() {
  final bd = ByteData(4);
  print(bd.lengthInBytes);
}
"#
        ),
        vec!["4"]
    );
}

#[test]
fn byte_data_view() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:typed_data';
void main() {
  final list = Uint8List.fromList([1, 2, 3, 4]);
  final bd = ByteData.view(list.buffer);
  print(bd.getUint8(1));
}
"#
        ),
        vec!["2"]
    );
}

#[test]
fn byte_data_view_offset() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:typed_data';
void main() {
  final list = Uint8List.fromList([10, 20, 30, 40]);
  final bd = ByteData.view(list.buffer, 2);
  print(bd.lengthInBytes);
  print(bd.getUint8(0));
}
"#
        ),
        vec!["2\n30"]
    );
}

#[test]
fn byte_data_get_int8() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:typed_data';
void main() {
  final bd = ByteData(1);
  bd.setInt8(0, -5);
  print(bd.getInt8(0));
}
"#
        ),
        vec!["-5"]
    );
}

#[test]
fn byte_data_get_uint8() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:typed_data';
void main() {
  final bd = ByteData(1);
  bd.setUint8(0, 250);
  print(bd.getUint8(0));
}
"#
        ),
        vec!["250"]
    );
}

#[test]
fn byte_data_get_int16_big_endian() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:typed_data';
void main() {
  final bd = ByteData(2);
  bd.setInt16(0, 0x1234, Endian.big);
  print(bd.getUint8(0) == 0x12);
  print(bd.getUint8(1) == 0x34);
}
"#
        ),
        vec!["true\ntrue"]
    );
}

#[test]
fn byte_data_get_int16_little_endian() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:typed_data';
void main() {
  final bd = ByteData(2);
  bd.setInt16(0, 0x1234, Endian.little);
  print(bd.getUint8(0) == 0x34);
  print(bd.getUint8(1) == 0x12);
}
"#
        ),
        vec!["true\ntrue"]
    );
}

#[test]
fn byte_data_get_uint16() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:typed_data';
void main() {
  final bd = ByteData(2);
  bd.setUint16(0, 65000);
  print(bd.getUint16(0));
}
"#
        ),
        vec!["65000"]
    );
}

#[test]
fn byte_data_get_int32() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:typed_data';
void main() {
  final bd = ByteData(4);
  bd.setInt32(0, -100000);
  print(bd.getInt32(0));
}
"#
        ),
        vec!["-100000"]
    );
}

#[test]
fn byte_data_get_uint32() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:typed_data';
void main() {
  final bd = ByteData(4);
  bd.setUint32(0, 4000000000);
  print(bd.getUint32(0));
}
"#
        ),
        vec!["4000000000"]
    );
}

#[test]
fn byte_data_get_int64() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:typed_data';
void main() {
  final bd = ByteData(8);
  bd.setInt64(0, -9000000000000000);
  print(bd.getInt64(0));
}
"#
        ),
        vec!["-9000000000000000"]
    );
}

#[test]
fn byte_data_get_uint64() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:typed_data';
void main() {
  final bd = ByteData(8);
  // Dart's JS integers are 53-bit, but native allows full 64-bit uint64 logic safely for typed data
  // We'll use a number < 2^53 for cross-platform safety just in case
  bd.setUint64(0, 9000000000000000);
  print(bd.getUint64(0));
}
"#
        ),
        vec!["9000000000000000"]
    );
}

#[test]
fn byte_data_get_float32() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:typed_data';
void main() {
  final bd = ByteData(4);
  bd.setFloat32(0, 3.125);
  print(bd.getFloat32(0));
}
"#
        ),
        vec!["3.125"]
    );
}

#[test]
fn byte_data_get_float64() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:typed_data';
void main() {
  final bd = ByteData(8);
  bd.setFloat64(0, 3.14159265359);
  print(bd.getFloat64(0));
}
"#
        ),
        vec!["3.14159265359"]
    );
}

#[test]
fn byte_data_out_of_bounds_throws() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:typed_data';
void main() {
  final bd = ByteData(2);
  try {
    bd.getInt32(0);
  } on RangeError {
    print('RangeError thrown');
  }
}
"#
        ),
        vec!["RangeError thrown"]
    );
}

#[test]
fn endian_host_property() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:typed_data';
void main() {
  print(Endian.host == Endian.little || Endian.host == Endian.big);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn byte_data_buffer_property() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:typed_data';
void main() {
  final bd = ByteData(4);
  print(bd.buffer is ByteBuffer);
  print(bd.buffer.lengthInBytes);
}
"#
        ),
        vec!["true\n4"]
    );
}

#[test]
fn byte_data_offset_in_bytes() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:typed_data';
void main() {
  final list = Uint8List(10);
  final bd = ByteData.view(list.buffer, 4);
  print(bd.offsetInBytes);
}
"#
        ),
        vec!["4"]
    );
}

#[test]
fn endianness_preservation_across_views() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:typed_data';
void main() {
  final bd = ByteData(4);
  bd.setUint32(0, 0xAABBCCDD, Endian.little);
  final u8 = Uint8List.view(bd.buffer);
  print('${u8[0].toRadixString(16).toUpperCase()}');
  // Little endian: Least significant byte at lowest address
}
"#
        ),
        vec!["DD"]
    );
}

#[test]
fn byte_data_set_float32_nan() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:typed_data';
void main() {
  final bd = ByteData(4);
  bd.setFloat32(0, double.nan);
  print(bd.getFloat32(0).isNaN);
}
"#
        ),
        vec!["true"]
    );
}
