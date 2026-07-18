use super::helpers::run_prints;

// ═══════════════════════════════════════════════════════════
// Dart: dart:typed_data Int Lists
// ═══════════════════════════════════════════════════════════

#[test]
fn int8list_creation_and_bounds() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:typed_data';
void main() {
  final l = Int8List(1);
  l[0] = 127;
  print(l[0]);
  l[0] = -128;
  print(l[0]);
}
"#
        ),
        vec!["127\n-128"]
    );
}

#[test]
fn int8list_truncation() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:typed_data';
void main() {
  final l = Int8List(1);
  l[0] = 128; // Truncated to -128
  print(l[0]);
}
"#
        ),
        vec!["-128"]
    );
}

#[test]
fn uint8list_creation_and_bounds() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:typed_data';
void main() {
  final l = Uint8List(1);
  l[0] = 255;
  print(l[0]);
}
"#
        ),
        vec!["255"]
    );
}

#[test]
fn uint8list_truncation() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:typed_data';
void main() {
  final l = Uint8List(1);
  l[0] = 256; // Truncated to 0
  print(l[0]);
  l[0] = -1; // Truncated to 255
  print(l[0]);
}
"#
        ),
        vec!["0\n255"]
    );
}

#[test]
fn int16list_creation_and_bounds() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:typed_data';
void main() {
  final l = Int16List(1);
  l[0] = 32767;
  print(l[0]);
  l[0] = -32768;
  print(l[0]);
}
"#
        ),
        vec!["32767\n-32768"]
    );
}

#[test]
fn int16list_truncation() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:typed_data';
void main() {
  final l = Int16List(1);
  l[0] = 32768; // Truncated to -32768
  print(l[0]);
}
"#
        ),
        vec!["-32768"]
    );
}

#[test]
fn uint16list_creation_and_bounds() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:typed_data';
void main() {
  final l = Uint16List(1);
  l[0] = 65535;
  print(l[0]);
}
"#
        ),
        vec!["65535"]
    );
}

#[test]
fn uint16list_truncation() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:typed_data';
void main() {
  final l = Uint16List(1);
  l[0] = 65536; // Truncated to 0
  print(l[0]);
  l[0] = -1; // Truncated to 65535
  print(l[0]);
}
"#
        ),
        vec!["0\n65535"]
    );
}

#[test]
fn int32list_creation() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:typed_data';
void main() {
  final l = Int32List(1);
  l[0] = 2147483647;
  print(l[0]);
}
"#
        ),
        vec!["2147483647"]
    );
}

#[test]
fn uint32list_creation() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:typed_data';
void main() {
  final l = Uint32List(1);
  l[0] = 4294967295;
  print(l[0]);
}
"#
        ),
        vec!["4294967295"]
    );
}

#[test]
fn int64list_creation() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:typed_data';
void main() {
  final l = Int64List(1);
  l[0] = -9000000000000000;
  print(l[0]);
}
"#
        ),
        vec!["-9000000000000000"]
    );
}

#[test]
fn uint64list_creation() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:typed_data';
void main() {
  final l = Uint64List(1);
  l[0] = 9000000000000000;
  print(l[0]);
}
"#
        ),
        vec!["9000000000000000"]
    );
}

#[test]
fn int32x4_creation() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:typed_data';
void main() {
  final l = Int32x4(1, 2, 3, 4);
  print('${l.x}:${l.y}:${l.z}:${l.w}');
}
"#
        ),
        vec!["1:2:3:4"]
    );
}

#[test]
fn int32x4_operations_add() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:typed_data';
void main() {
  final a = Int32x4(10, 20, 30, 40);
  final b = Int32x4(1, 2, 3, 4);
  final c = a + b;
  print(c.y);
}
"#
        ),
        vec!["22"]
    );
}

#[test]
fn int32x4_operations_bitwise_and() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:typed_data';
void main() {
  final a = Int32x4(15, 15, 15, 15);
  final b = Int32x4(3, 3, 3, 3);
  final c = a & b;
  print(c.x);
}
"#
        ),
        vec!["3"]
    );
}

#[test]
fn int32x4list_creation() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:typed_data';
void main() {
  final l = Int32x4List(1);
  l[0] = Int32x4(9, 8, 7, 6);
  print(l[0].z);
}
"#
        ),
        vec!["7"]
    );
}

#[test]
fn int_list_sublist() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:typed_data';
void main() {
  final l = Uint8List.fromList([10, 20, 30, 40]);
  final sub = l.sublist(1, 3);
  print(sub is Uint8List);
  print(sub.length);
  print(sub[0]);
}
"#
        ),
        vec!["true\n2\n20"]
    );
}

#[test]
fn int_list_set_range() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:typed_data';
void main() {
  final l1 = Uint8List(5);
  final l2 = Uint8List.fromList([1, 2, 3]);
  l1.setRange(1, 4, l2);
  print('${l1[0]}:${l1[1]}:${l1[3]}:${l1[4]}');
}
"#
        ),
        vec!["0:1:3:0"]
    );
}

#[test]
fn int_list_fill_range() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:typed_data';
void main() {
  final l = Int16List(4);
  l.fillRange(1, 3, 99);
  print('${l[0]}:${l[1]}:${l[2]}:${l[3]}');
}
"#
        ),
        vec!["0:99:99:0"]
    );
}

#[test]
fn uint8clampedlist_clamping() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:typed_data';
void main() {
  final l = Uint8ClampedList(2);
  l[0] = 300; // Clamped to 255
  l[1] = -50; // Clamped to 0
  print('${l[0]}:${l[1]}');
}
"#
        ),
        vec!["255:0"]
    );
}
