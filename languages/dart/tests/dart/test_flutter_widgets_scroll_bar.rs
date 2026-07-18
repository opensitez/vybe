use super::helpers::run_prints;

// ═══════════════════════════════════════════════════════════
// Flutter: widgets Scrollbar
// ═══════════════════════════════════════════════════════════

#[test]
fn scrollbar_creation() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  const sb = Scrollbar(
    child: SizedBox(),
  );
  print(sb is StatelessWidget);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn scrollbar_child() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  const sb = Scrollbar(
    child: Text('Scroll'),
  );
  print(sb.child is Text);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn scrollbar_controller() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final sc = ScrollController();
  final sb = Scrollbar(
    controller: sc,
    child: const SizedBox(),
  );
  print(sb.controller == sc);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn scrollbar_thumb_visibility() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  const sb = Scrollbar(
    thumbVisibility: true,
    child: SizedBox(),
  );
  print(sb.thumbVisibility);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn scrollbar_track_visibility() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  const sb = Scrollbar(
    trackVisibility: true,
    child: SizedBox(),
  );
  print(sb.trackVisibility);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn scrollbar_thickness() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  const sb = Scrollbar(
    thickness: 10.0,
    child: SizedBox(),
  );
  print(sb.thickness);
}
"#
        ),
        vec!["10.0"]
    );
}

#[test]
fn scrollbar_radius() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  const sb = Scrollbar(
    radius: Radius.circular(5.0),
    child: SizedBox(),
  );
  print((sb.radius as Radius).x);
}
"#
        ),
        vec!["5.0"]
    );
}
