use super::helpers::run_prints;

// ═══════════════════════════════════════════════════════════
// Dart: dart:typed_data Float Lists
// ═══════════════════════════════════════════════════════════

#[test]
fn float32list_creation() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:typed_data';
void main() {
  final f = Float32List(3);
  print(f.length);
  print(f[0]);
}
"#
        ),
        vec!["3\n0.0"]
    );
}

#[test]
fn float32list_from_list() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:typed_data';
void main() {
  final f = Float32List.fromList([1.5, 2.25, 3.125]);
  print(f[1]);
}
"#
        ),
        vec!["2.25"]
    );
}

#[test]
fn float32list_precision_loss() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:typed_data';
void main() {
  // Float32 loses precision on double
  final f = Float32List(1);
  f[0] = 3.141592653589793238; // double precision
  print(f[0] == 3.141592653589793238); // false
  print(f[0] != 0.0);
}
"#
        ),
        vec!["false\ntrue"]
    );
}

#[test]
fn float64list_creation() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:typed_data';
void main() {
  final f = Float64List(2);
  print(f.length);
  print(f[1]);
}
"#
        ),
        vec!["2\n0.0"]
    );
}

#[test]
fn float64list_from_list() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:typed_data';
void main() {
  final f = Float64List.fromList([1.1, 2.2]);
  print(f[0]);
}
"#
        ),
        vec!["1.1"]
    );
}

#[test]
fn float64list_maintains_precision() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:typed_data';
void main() {
  final f = Float64List(1);
  double val = 3.141592653589793;
  f[0] = val;
  print(f[0] == val);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn float32list_nan() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:typed_data';
void main() {
  final f = Float32List(1);
  f[0] = double.nan;
  print(f[0].isNaN);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn float64list_nan() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:typed_data';
void main() {
  final f = Float64List(1);
  f[0] = double.nan;
  print(f[0].isNaN);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn float32list_infinity() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:typed_data';
void main() {
  final f = Float32List(2);
  f[0] = double.infinity;
  f[1] = double.negativeInfinity;
  print(f[0] == double.infinity);
  print(f[1] == double.negativeInfinity);
}
"#
        ),
        vec!["true\ntrue"]
    );
}

#[test]
fn float32x4_creation() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:typed_data';
void main() {
  final f = Float32x4(1.0, 2.0, 3.0, 4.0);
  print('${f.x}:${f.y}:${f.z}:${f.w}');
}
"#
        ),
        vec!["1.0:2.0:3.0:4.0"]
    );
}

#[test]
fn float32x4_splat() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:typed_data';
void main() {
  final f = Float32x4.splat(5.0);
  print('${f.x}:${f.w}');
}
"#
        ),
        vec!["5.0:5.0"]
    );
}

#[test]
fn float32x4_zero() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:typed_data';
void main() {
  final f = Float32x4.zero();
  print(f.x == 0.0);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn float32x4_operations_add() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:typed_data';
void main() {
  final a = Float32x4(1.0, 2.0, 3.0, 4.0);
  final b = Float32x4(10.0, 20.0, 30.0, 40.0);
  final c = a + b;
  print(c.y);
}
"#
        ),
        vec!["22.0"]
    );
}

#[test]
fn float32x4_operations_mul() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:typed_data';
void main() {
  final a = Float32x4(2.0, 2.0, 2.0, 2.0);
  final b = Float32x4(3.0, 4.0, 5.0, 6.0);
  final c = a * b;
  print(c.w);
}
"#
        ),
        vec!["12.0"]
    );
}

#[test]
fn float32x4list_creation() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:typed_data';
void main() {
  final l = Float32x4List(2);
  print(l.length);
  print(l[0].x);
}
"#
        ),
        vec!["2\n0.0"]
    );
}

#[test]
fn float32x4list_from_list() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:typed_data';
void main() {
  final l = Float32x4List.fromList([Float32x4(1.0, 2.0, 3.0, 4.0)]);
  print(l[0].z);
}
"#
        ),
        vec!["3.0"]
    );
}

#[test]
fn float64x2_creation() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:typed_data';
void main() {
  final f = Float64x2(1.5, 2.5);
  print('${f.x}:${f.y}');
}
"#
        ),
        vec!["1.5:2.5"]
    );
}

#[test]
fn float64x2_operations_sub() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:typed_data';
void main() {
  final a = Float64x2(5.0, 10.0);
  final b = Float64x2(2.0, 3.0);
  final c = a - b;
  print(c.x);
}
"#
        ),
        vec!["3.0"]
    );
}

#[test]
fn float64x2list_creation() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:typed_data';
void main() {
  final l = Float64x2List(1);
  l[0] = Float64x2(9.9, 8.8);
  print(l[0].y);
}
"#
        ),
        vec!["8.8"]
    );
}

#[test]
fn float_list_view_byte_data() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:typed_data';
void main() {
  final list = Float32List(1);
  list[0] = 1.0; // 0x3F800000
  final bd = ByteData.view(list.buffer);
  print(bd.getFloat32(0, Endian.host));
}
"#
        ),
        vec!["1.0"]
    );
}
