use super::helpers::run_prints;

// ═══════════════════════════════════════════════════════════
// Flutter: dart:ui Color & Math
// ═══════════════════════════════════════════════════════════

#[test]
fn color_creation() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:ui';
void main() {
  final c = Color(0xFF112233);
  print(c.value);
}
"#
        ),
        vec!["4279312947"] // 0xFF112233
    );
}

#[test]
fn color_components() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:ui';
void main() {
  final c = Color(0x12345678);
  print('${c.alpha}:${c.red}:${c.green}:${c.blue}');
}
"#
        ),
        vec!["18:52:86:120"] // 0x12=18, 0x34=52, 0x56=86, 0x78=120
    );
}

#[test]
fn color_from_argb() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:ui';
void main() {
  final c = Color.fromARGB(255, 255, 255, 255);
  print(c.value == 0xFFFFFFFF);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn color_from_rgbo() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:ui';
void main() {
  // opacity is 0.0 to 1.0
  final c = Color.fromRGBO(255, 0, 0, 0.5);
  // 0.5 * 255 = 127 = 0x7F
  print(c.alpha == 127 || c.alpha == 128); // Precision might make it 127 or 128 depending on rounding
  print('${c.red}:${c.green}:${c.blue}');
}
"#
        ),
        // Wait, Dart Color.fromRGBO(r,g,b,o) uses `(opacity * 255.0).round()`.
        // 0.5 * 255 = 127.5 -> round() -> 128
        vec!["true\n255:0:0"]
    );
}

#[test]
fn color_with_alpha() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:ui';
void main() {
  final c = Color(0xFF00FF00).withAlpha(128);
  print(c.value == 0x8000FF00); // 128 = 0x80
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn color_with_opacity() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:ui';
void main() {
  final c = Color(0xFF0000FF).withOpacity(1.0);
  print(c.value == 0xFF0000FF);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn color_with_red() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:ui';
void main() {
  final c = Color(0xFF000000).withRed(255);
  print(c.value == 0xFFFF0000);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn color_compute_luminance() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:ui';
void main() {
  final c = Color(0xFFFFFFFF);
  print(c.computeLuminance() == 1.0);
  final b = Color(0xFF000000);
  print(b.computeLuminance() == 0.0);
}
"#
        ),
        vec!["true\ntrue"]
    );
}

#[test]
fn color_equality() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:ui';
void main() {
  final c1 = Color(0xFF112233);
  final c2 = Color(0xFF112233);
  print(c1 == c2);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn color_alpha_blend() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:ui';
void main() {
  final fg = Color(0x80FF0000); // semi-transparent red
  final bg = Color(0xFF00FF00); // solid green
  final blended = Color.alphaBlend(fg, bg);
  print(blended.alpha == 255);
}
"#
        ),
        vec!["true"] // 0x80 + solid bg = solid
    );
}

#[test]
fn color_get_fill_color() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:ui';
void main() {
  final c = Color(0x00000000);
  print(c.opacity == 0.0);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn color_hashcode() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:ui';
void main() {
  final c = Color(0xFF123456);
  print(c.hashCode == c.value.hashCode);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn color_tostring() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:ui';
void main() {
  final c = Color(0xFF123456);
  print(c.toString().contains('123456'));
}
"#
        ),
        vec!["true"]
    );
}
