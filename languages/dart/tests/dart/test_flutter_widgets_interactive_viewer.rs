use super::helpers::run_prints;

// ═══════════════════════════════════════════════════════════
// Flutter: widgets InteractiveViewer
// ═══════════════════════════════════════════════════════════

#[test]
fn interactive_viewer_creation() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final iv = InteractiveViewer(child: const SizedBox());
  print(iv is StatefulWidget);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn interactive_viewer_clip_behavior() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final iv = InteractiveViewer(
    clipBehavior: Clip.none,
    child: const SizedBox(),
  );
  print(iv.clipBehavior == Clip.none);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn interactive_viewer_pan_enabled() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final iv = InteractiveViewer(
    panEnabled: false,
    child: const SizedBox(),
  );
  print(iv.panEnabled);
}
"#
        ),
        vec!["false"]
    );
}

#[test]
fn interactive_viewer_scale_enabled() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final iv = InteractiveViewer(
    scaleEnabled: false,
    child: const SizedBox(),
  );
  print(iv.scaleEnabled);
}
"#
        ),
        vec!["false"]
    );
}

#[test]
fn interactive_viewer_min_scale() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final iv = InteractiveViewer(
    minScale: 0.5,
    child: const SizedBox(),
  );
  print(iv.minScale);
}
"#
        ),
        vec!["0.5"]
    );
}

#[test]
fn interactive_viewer_max_scale() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final iv = InteractiveViewer(
    maxScale: 4.0,
    child: const SizedBox(),
  );
  print(iv.maxScale);
}
"#
        ),
        vec!["4.0"]
    );
}

#[test]
fn interactive_viewer_constrained() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final iv = InteractiveViewer(
    constrained: false,
    child: const SizedBox(),
  );
  print(iv.constrained);
}
"#
        ),
        vec!["false"]
    );
}

#[test]
fn interactive_viewer_boundary_margin() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final iv = InteractiveViewer(
    boundaryMargin: const EdgeInsets.all(10.0),
    child: const SizedBox(),
  );
  print((iv.boundaryMargin as EdgeInsets).top);
}
"#
        ),
        vec!["10.0"]
    );
}
