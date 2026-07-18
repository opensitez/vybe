use super::helpers::run_prints;

// ═══════════════════════════════════════════════════════════
// Flutter: widgets Image
// ═══════════════════════════════════════════════════════════

#[test]
fn image_creation() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
import 'dart:typed_data';
void main() {
  final img = Image.memory(Uint8List(0));
  print(img != null);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn image_network() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final img = Image.network('https://example.com/img.png');
  print(img.image is NetworkImage);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn image_asset() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final img = Image.asset('assets/logo.png');
  print(img.image is AssetImage);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn image_dimensions() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
import 'dart:typed_data';
void main() {
  final img = Image.memory(Uint8List(0), width: 100, height: 200);
  print('${img.width}:${img.height}');
}
"#
        ),
        vec!["100.0:200.0"]
    );
}

#[test]
fn image_fit() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
import 'dart:typed_data';
void main() {
  final img = Image.memory(Uint8List(0), fit: BoxFit.cover);
  print(img.fit == BoxFit.cover);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn image_alignment() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
import 'dart:typed_data';
void main() {
  final img = Image.memory(Uint8List(0), alignment: Alignment.bottomLeft);
  print(img.alignment == Alignment.bottomLeft);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn image_repeat() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
import 'dart:typed_data';
void main() {
  final img = Image.memory(Uint8List(0), repeat: ImageRepeat.repeatX);
  print(img.repeat == ImageRepeat.repeatX);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn image_color_blend_mode() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
import 'dart:typed_data';
void main() {
  final img = Image.memory(Uint8List(0), color: const Color(0xFFFF0000), colorBlendMode: BlendMode.srcOver);
  print(img.colorBlendMode == BlendMode.srcOver);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn image_filter_quality() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
import 'dart:typed_data';
void main() {
  final img = Image.memory(Uint8List(0), filterQuality: FilterQuality.high);
  print(img.filterQuality == FilterQuality.high);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn image_is_stateful() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
import 'dart:typed_data';
void main() {
  final img = Image.memory(Uint8List(0));
  print(img is StatefulWidget);
}
"#
        ),
        vec!["true"]
    );
}
