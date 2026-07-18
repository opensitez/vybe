use super::helpers::run_prints;

// ═══════════════════════════════════════════════════════════
// Dart: dart:typed_data Unmodifiable Views
// ═══════════════════════════════════════════════════════════

#[test]
fn unmodifiable_uint8list_view_creation() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:typed_data';
void main() {
  final l = Uint8List.fromList([1, 2, 3]);
  final ul = UnmodifiableUint8ListView(l);
  print(ul[1]);
}
"#
        ),
        vec!["2"]
    );
}

#[test]
fn unmodifiable_uint8list_view_mutation_throws() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:typed_data';
void main() {
  final l = Uint8List.fromList([1, 2, 3]);
  final ul = UnmodifiableUint8ListView(l);
  try {
    ul[0] = 10;
  } on UnsupportedError {
    print('UnsupportedError thrown');
  }
}
"#
        ),
        vec!["UnsupportedError thrown"]
    );
}

#[test]
fn unmodifiable_int8list_view_creation() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:typed_data';
void main() {
  final l = Int8List.fromList([-1, -2]);
  final ul = UnmodifiableInt8ListView(l);
  print(ul[0]);
}
"#
        ),
        vec!["-1"]
    );
}

#[test]
fn unmodifiable_uint16list_view_creation() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:typed_data';
void main() {
  final l = Uint16List.fromList([1000, 2000]);
  final ul = UnmodifiableUint16ListView(l);
  print(ul[1]);
}
"#
        ),
        vec!["2000"]
    );
}

#[test]
fn unmodifiable_int16list_view_mutation_throws() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:typed_data';
void main() {
  final l = Int16List.fromList([-1000]);
  final ul = UnmodifiableInt16ListView(l);
  try {
    ul[0] = 0;
  } on UnsupportedError {
    print('UnsupportedError thrown');
  }
}
"#
        ),
        vec!["UnsupportedError thrown"]
    );
}

#[test]
fn unmodifiable_uint32list_view_creation() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:typed_data';
void main() {
  final l = Uint32List.fromList([4000000000]);
  final ul = UnmodifiableUint32ListView(l);
  print(ul[0]);
}
"#
        ),
        vec!["4000000000"]
    );
}

#[test]
fn unmodifiable_int32list_view_creation() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:typed_data';
void main() {
  final l = Int32List.fromList([-2000000000]);
  final ul = UnmodifiableInt32ListView(l);
  print(ul[0]);
}
"#
        ),
        vec!["-2000000000"]
    );
}

#[test]
fn unmodifiable_uint64list_view_creation() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:typed_data';
void main() {
  final l = Uint64List.fromList([9000000000000000]);
  final ul = UnmodifiableUint64ListView(l);
  print(ul[0]);
}
"#
        ),
        vec!["9000000000000000"]
    );
}

#[test]
fn unmodifiable_int64list_view_mutation_throws() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:typed_data';
void main() {
  final l = Int64List.fromList([1]);
  final ul = UnmodifiableInt64ListView(l);
  try {
    ul[0] = 2;
  } on UnsupportedError {
    print('UnsupportedError thrown');
  }
}
"#
        ),
        vec!["UnsupportedError thrown"]
    );
}

#[test]
fn unmodifiable_float32list_view_creation() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:typed_data';
void main() {
  final l = Float32List.fromList([3.5]);
  final ul = UnmodifiableFloat32ListView(l);
  print(ul[0]);
}
"#
        ),
        vec!["3.5"]
    );
}

#[test]
fn unmodifiable_float64list_view_mutation_throws() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:typed_data';
void main() {
  final l = Float64List.fromList([1.1]);
  final ul = UnmodifiableFloat64ListView(l);
  try {
    ul[0] = 2.2;
  } on UnsupportedError {
    print('UnsupportedError thrown');
  }
}
"#
        ),
        vec!["UnsupportedError thrown"]
    );
}

#[test]
fn unmodifiable_float32x4list_view_creation() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:typed_data';
void main() {
  final l = Float32x4List.fromList([Float32x4(1.0, 2.0, 3.0, 4.0)]);
  final ul = UnmodifiableFloat32x4ListView(l);
  print(ul[0].z);
}
"#
        ),
        vec!["3.0"]
    );
}

#[test]
fn unmodifiable_float64x2list_view_creation() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:typed_data';
void main() {
  final l = Float64x2List.fromList([Float64x2(10.0, 20.0)]);
  final ul = UnmodifiableFloat64x2ListView(l);
  print(ul[0].y);
}
"#
        ),
        vec!["20.0"]
    );
}

#[test]
fn unmodifiable_int32x4list_view_creation() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:typed_data';
void main() {
  final l = Int32x4List.fromList([Int32x4(1, 2, 3, 4)]);
  final ul = UnmodifiableInt32x4ListView(l);
  print(ul[0].w);
}
"#
        ),
        vec!["4"]
    );
}

#[test]
fn unmodifiable_byte_data_view_creation() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:typed_data';
void main() {
  final l = Uint8List.fromList([1, 2, 3, 4]);
  final bd = ByteData.view(l.buffer);
  final ubd = UnmodifiableByteDataView(bd);
  print(ubd.getUint8(1));
}
"#
        ),
        vec!["2"]
    );
}

#[test]
fn unmodifiable_byte_data_view_mutation_throws() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:typed_data';
void main() {
  final bd = ByteData(4);
  final ubd = UnmodifiableByteDataView(bd);
  try {
    ubd.setUint8(0, 10);
  } on UnsupportedError {
    print('UnsupportedError thrown');
  }
}
"#
        ),
        vec!["UnsupportedError thrown"]
    );
}

#[test]
fn unmodifiable_byte_data_view_buffer_property() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:typed_data';
void main() {
  final bd = ByteData(4);
  final ubd = UnmodifiableByteDataView(bd);
  print(ubd.buffer is ByteBuffer);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn unmodifiable_uint8clampedlist_view_creation() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:typed_data';
void main() {
  final l = Uint8ClampedList.fromList([10, 20]);
  final ul = UnmodifiableUint8ClampedListView(l);
  print(ul[1]);
}
"#
        ),
        vec!["20"]
    );
}

#[test]
fn unmodifiable_uint8list_view_reflects_original() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:typed_data';
void main() {
  final l = Uint8List(1);
  final ul = UnmodifiableUint8ListView(l);
  l[0] = 99;
  print(ul[0]);
}
"#
        ),
        vec!["99"]
    );
}

#[test]
fn unmodifiable_byte_data_view_reflects_original() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:typed_data';
void main() {
  final bd = ByteData(4);
  final ubd = UnmodifiableByteDataView(bd);
  bd.setInt32(0, 12345);
  print(ubd.getInt32(0));
}
"#
        ),
        vec!["12345"]
    );
}
