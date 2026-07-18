use super::helpers::run_prints;

// ═══════════════════════════════════════════════════════════
// Flutter: widgets PageRoute
// ═══════════════════════════════════════════════════════════

#[test]
fn page_route_builder_creation() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final r = PageRouteBuilder(pageBuilder: (c, a, sa) => const SizedBox());
  print(r.barrierColor == null);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn page_route_builder_opaque() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final r = PageRouteBuilder(pageBuilder: (c, a, sa) => const SizedBox(), opaque: false);
  print(r.opaque);
}
"#
        ),
        vec!["false"]
    );
}

#[test]
fn page_route_builder_barrier_dismissible() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final r = PageRouteBuilder(pageBuilder: (c, a, sa) => const SizedBox(), barrierDismissible: true);
  print(r.barrierDismissible);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn page_route_builder_maintain_state() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final r = PageRouteBuilder(pageBuilder: (c, a, sa) => const SizedBox(), maintainState: false);
  print(r.maintainState);
}
"#
        ),
        vec!["false"]
    );
}

#[test]
fn page_route_builder_fullscreen_dialog() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final r = PageRouteBuilder(pageBuilder: (c, a, sa) => const SizedBox(), fullscreenDialog: true);
  print(r.fullscreenDialog);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn modal_route_of() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final w = const SizedBox();
  final e = w.createElement();
  final r = ModalRoute.of(e);
  print(r == null);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn local_history_route() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final r = PageRouteBuilder(pageBuilder: (c, a, sa) => const SizedBox());
  print(r is LocalHistoryRoute);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn page_route_transition_duration() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final r = PageRouteBuilder(
    pageBuilder: (c, a, sa) => const SizedBox(),
    transitionDuration: Duration(milliseconds: 300)
  );
  print(r.transitionDuration.inMilliseconds);
}
"#
        ),
        vec!["300"]
    );
}

#[test]
fn page_route_reverse_transition_duration() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final r = PageRouteBuilder(
    pageBuilder: (c, a, sa) => const SizedBox(),
    reverseTransitionDuration: Duration(milliseconds: 200)
  );
  print(r.reverseTransitionDuration.inMilliseconds);
}
"#
        ),
        vec!["200"]
    );
}

#[test]
fn page_route_can_transition_to() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final r1 = PageRouteBuilder(pageBuilder: (c, a, sa) => const SizedBox());
  final r2 = PageRouteBuilder(pageBuilder: (c, a, sa) => const SizedBox());
  print(r1.canTransitionTo(r2));
}
"#
        ),
        vec!["true"]
    );
}
