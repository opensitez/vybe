use super::helpers::run_prints;

// ═══════════════════════════════════════════════════════════
// Dart: dart:typed_data Buffer Views
// ═══════════════════════════════════════════════════════════

#[test]
fn uint8list_view() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:typed_data';
void main() {
  final buffer = Uint8List.fromList([1, 2, 3, 4]).buffer;
  final view = Uint8List.view(buffer, 1, 2);
  print('${view.length}:${view[0]}');
}
"#
        ),
        vec!["2\n2"]
    );
}

#[test]
fn int8list_view() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:typed_data';
void main() {
  final buffer = Uint8List.fromList([255, 10, 20]).buffer;
  final view = Int8List.view(buffer);
  print(view[0]); // 255 in Int8 is -1
}
"#
        ),
        vec!["-1"]
    );
}

#[test]
fn uint16list_view() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:typed_data';
void main() {
  final bd = ByteData(4);
  bd.setUint16(0, 500, Endian.host);
  bd.setUint16(2, 1000, Endian.host);
  final view = Uint16List.view(bd.buffer);
  print('${view.length}:${view[0]}:${view[1]}');
}
"#
        ),
        vec!["2\n500\n1000"]
    );
}

#[test]
fn uint16list_view_invalid_offset_throws() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:typed_data';
void main() {
  final bd = ByteData(4);
  try {
    // Offset must be a multiple of 2 for Uint16List
    Uint16List.view(bd.buffer, 1);
  } on ArgumentError {
    print('ArgumentError thrown');
  }
}
"#
        ),
        vec!["ArgumentError thrown"]
    );
}

#[test]
fn int32list_view() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:typed_data';
void main() {
  final bd = ByteData(8);
  bd.setInt32(4, -123456, Endian.host);
  final view = Int32List.view(bd.buffer, 4, 1);
  print(view[0]);
}
"#
        ),
        vec!["-123456"]
    );
}

#[test]
fn int32list_view_invalid_offset_throws() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:typed_data';
void main() {
  final bd = ByteData(8);
  try {
    Int32List.view(bd.buffer, 2);
  } on ArgumentError {
    print('ArgumentError thrown');
  }
}
"#
        ),
        vec!["ArgumentError thrown"]
    );
}

#[test]
fn float32list_view() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:typed_data';
void main() {
  final bd = ByteData(8);
  bd.setFloat32(0, 3.5, Endian.host);
  final view = Float32List.view(bd.buffer, 0, 1);
  print(view[0]);
}
"#
        ),
        vec!["3.5"]
    );
}

#[test]
fn float64list_view() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:typed_data';
void main() {
  final bd = ByteData(16);
  bd.setFloat64(8, 3.14, Endian.host);
  final view = Float64List.view(bd.buffer, 8, 1);
  print(view[0]);
}
"#
        ),
        vec!["3.14"]
    );
}

#[test]
fn uint8clampedlist_view() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:typed_data';
void main() {
  final buffer = Uint8List.fromList([10, 255]).buffer;
  final view = Uint8ClampedList.view(buffer, 1, 1);
  print(view[0]);
}
"#
        ),
        vec!["255"]
    );
}

#[test]
fn buffer_shared_memory() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:typed_data';
void main() {
  final list1 = Uint8List(4);
  final list2 = Uint16List.view(list1.buffer);
  list1[0] = 0xFF;
  list1[1] = 0x00;
  // If host is little endian, 0x00FF = 255
  // If host is big endian, 0xFF00 = 65280
  final v = list2[0];
  print(v == 255 || v == 65280);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn view_offset_in_bytes() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:typed_data';
void main() {
  final list = Uint8List(10);
  final view = Uint32List.view(list.buffer, 4, 1);
  print(view.offsetInBytes);
}
"#
        ),
        vec!["4"]
    );
}

#[test]
fn view_length_in_bytes() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:typed_data';
void main() {
  final list = Uint8List(16);
  final view = Float64List.view(list.buffer, 0, 2);
  print(view.lengthInBytes);
}
"#
        ),
        vec!["16"]
    );
}

#[test]
fn float32x4list_view() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:typed_data';
void main() {
  final bd = ByteData(16);
  bd.setFloat32(0, 1.0, Endian.host);
  bd.setFloat32(4, 2.0, Endian.host);
  bd.setFloat32(8, 3.0, Endian.host);
  bd.setFloat32(12, 4.0, Endian.host);
  final view = Float32x4List.view(bd.buffer);
  print(view[0].z);
}
"#
        ),
        vec!["3.0"]
    );
}

#[test]
fn float64x2list_view() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:typed_data';
void main() {
  final bd = ByteData(16);
  bd.setFloat64(0, 10.0, Endian.host);
  bd.setFloat64(8, 20.0, Endian.host);
  final view = Float64x2List.view(bd.buffer);
  print('${view[0].x}:${view[0].y}');
}
"#
        ),
        vec!["10.0:20.0"]
    );
}

#[test]
fn int32x4list_view() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:typed_data';
void main() {
  final bd = ByteData(16);
  bd.setInt32(0, 1, Endian.host);
  bd.setInt32(4, 2, Endian.host);
  bd.setInt32(8, 3, Endian.host);
  bd.setInt32(12, 4, Endian.host);
  final view = Int32x4List.view(bd.buffer);
  print(view[0].w);
}
"#
        ),
        vec!["4"]
    );
}

#[test]
fn uint64list_view() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:typed_data';
void main() {
  final bd = ByteData(8);
  // Dart might truncate > 53 bit numbers in JSON, but typed_data should hold 64 bits natively.
  bd.setUint64(0, 123456789, Endian.host);
  final view = Uint64List.view(bd.buffer);
  print(view[0]);
}
"#
        ),
        vec!["123456789"]
    );
}

#[test]
fn int64list_view() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:typed_data';
void main() {
  final bd = ByteData(8);
  bd.setInt64(0, -123456789, Endian.host);
  final view = Int64List.view(bd.buffer);
  print(view[0]);
}
"#
        ),
        vec!["-123456789"]
    );
}

#[test]
fn int16list_view_out_of_bounds_throws() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:typed_data';
void main() {
  final bd = ByteData(3);
  try {
    Int16List.view(bd.buffer, 2);
  } on ArgumentError {
    print('ArgumentError thrown');
  }
}
"#
        ),
        vec!["ArgumentError thrown"]
    );
}

#[test]
fn byte_data_view_bounds_check() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:typed_data';
void main() {
  final buffer = Uint8List(5).buffer;
  try {
    ByteData.view(buffer, 6);
  } on RangeError {
    print('RangeError thrown');
  } catch(e) {
    print('ArgumentError thrown'); // Depending on impl
  }
}
"#
        ),
        vec!["RangeError thrown"] // Dart specifically throws RangeError for invalid offset in ByteData.view
    );
}

#[test]
fn buffer_view_modifies_original() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:typed_data';
void main() {
  final list = Uint8List(4);
  final view = Uint32List.view(list.buffer);
  view[0] = 0xFFFFFFFF;
  print('${list[0]}:${list[1]}:${list[2]}:${list[3]}');
}
"#
        ),
        vec!["255:255:255:255"]
    );
}
