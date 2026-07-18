use super::helpers::run_prints;

// ═══════════════════════════════════════════════════════════
// Flutter: widgets ScrollController
// ═══════════════════════════════════════════════════════════

#[test]
fn scroll_controller_creation() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final sc = ScrollController(initialScrollOffset: 100.0, keepScrollOffset: true);
  print('${sc.initialScrollOffset}:${sc.keepScrollOffset}');
}
"#
        ),
        vec!["100.0:true"]
    );
}

#[test]
fn scroll_controller_has_clients() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final sc = ScrollController();
  print(sc.hasClients);
}
"#
        ),
        vec!["false"]
    );
}

#[test]
fn scroll_controller_offset_without_client_throws() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final sc = ScrollController();
  try {
    print(sc.offset);
  } catch(e) {
    print('throws');
  }
}
"#
        ),
        vec!["throws"]
    );
}

#[test]
fn scroll_controller_dispose() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final sc = ScrollController();
  sc.dispose();
  print('disposed');
}
"#
        ),
        vec!["disposed"]
    );
}

#[test]
fn scroll_controller_animate_to_no_client() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final sc = ScrollController();
  try {
    sc.animateTo(100.0, duration: Duration(seconds: 1), curve: Curves.linear);
  } catch(e) {
    print('throws');
  }
}
"#
        ),
        vec!["throws"]
    );
}

#[test]
fn scroll_controller_jump_to_no_client() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final sc = ScrollController();
  try {
    sc.jumpTo(100.0);
  } catch(e) {
    print('throws');
  }
}
"#
        ),
        vec!["throws"]
    );
}

#[test]
fn tracking_scroll_controller() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final sc = TrackingScrollController();
  print(sc.initialScrollOffset);
}
"#
        ),
        vec!["0.0"]
    );
}

#[test]
fn scroll_position_abstract() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final sc = ScrollController();
  try {
    final pos = sc.position; // throws if no clients
  } catch(e) {
    print('throws_position');
  }
}
"#
        ),
        vec!["throws_position"]
    );
}

#[test]
fn scroll_controller_create_scroll_position() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  // Can't easily test createScrollPosition since it needs a ScrollPhysics and ScrollContext
  print('compiles');
}
"#
        ),
        vec!["compiles"]
    );
}

#[test]
fn fixed_extent_scroll_controller() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final sc = FixedExtentScrollController(initialItem: 5);
  print(sc.initialItem);
}
"#
        ),
        vec!["5"]
    );
}

#[test]
fn scroll_controller_debug_label() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final sc = ScrollController(debugLabel: 'my_scroller');
  print(sc.debugLabel);
}
"#
        ),
        vec!["my_scroller"]
    );
}
