use super::helpers::run_prints;

// ═══════════════════════════════════════════════════════════
// Flutter: widgets ImageProvider
// ═══════════════════════════════════════════════════════════

#[test]
fn exact_asset_image_creation() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final img = ExactAssetImage('assets/test.png');
  print(img.assetName);
}
"#
        ),
        vec!["assets/test.png"]
    );
}

#[test]
fn exact_asset_image_scale() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final img = ExactAssetImage('assets/test.png', scale: 2.0);
  print(img.scale);
}
"#
        ),
        vec!["2.0"]
    );
}

#[test]
fn network_image_creation() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final img = NetworkImage('https://example.com/test.png');
  print(img.url);
}
"#
        ),
        vec!["https://example.com/test.png"]
    );
}

#[test]
fn network_image_scale() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final img = NetworkImage('https://example.com/test.png', scale: 1.5);
  print(img.scale);
}
"#
        ),
        vec!["1.5"]
    );
}

#[test]
fn file_image_creation() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
import 'dart:io';
void main() {
  final file = File('/tmp/test.png');
  final img = FileImage(file);
  print(img.file.path);
}
"#
        ),
        vec!["/tmp/test.png"]
    );
}

#[test]
fn memory_image_creation() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
import 'dart:typed_data';
void main() {
  final bytes = Uint8List.fromList([1, 2, 3]);
  final img = MemoryImage(bytes);
  print(img.bytes.length);
}
"#
        ),
        vec!["3"]
    );
}

#[test]
fn resize_image_creation() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final base = NetworkImage('https://example.com/test.png');
  final resized = ResizeImage(base, width: 100, height: 100);
  print('${resized.width}:${resized.height}');
}
"#
        ),
        vec!["100:100"]
    );
}

#[test]
fn image_provider_resolve() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final img = NetworkImage('https://example.com/test.png');
  final stream = img.resolve(ImageConfiguration.empty);
  print(stream != null);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn image_configuration_empty() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final c = ImageConfiguration.empty;
  print(c.size == null);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn image_configuration_properties() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
import 'dart:ui';
void main() {
  final c = ImageConfiguration(size: Size(100, 100), devicePixelRatio: 2.0);
  print('${c.size!.width}:${c.devicePixelRatio}');
}
"#
        ),
        vec!["100.0:2.0"]
    );
}
