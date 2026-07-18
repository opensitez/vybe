use super::helpers::run_prints;

// ═══════════════════════════════════════════════════════════
// Flutter: widgets PageView
// ═══════════════════════════════════════════════════════════

#[test]
fn page_view_creation() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final pv = PageView(children: const []);
  print(pv is StatefulWidget);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn page_view_scroll_direction() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final pv = PageView(
    scrollDirection: Axis.vertical,
    children: const [],
  );
  print(pv.scrollDirection == Axis.vertical);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn page_view_reverse() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final pv = PageView(
    reverse: true,
    children: const [],
  );
  print(pv.reverse);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn page_view_controller() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final pc = PageController(initialPage: 2);
  final pv = PageView(
    controller: pc,
    children: const [],
  );
  print((pv.controller as PageController).initialPage);
}
"#
        ),
        vec!["2"]
    );
}

#[test]
fn page_view_physics() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final pv = PageView(
    physics: const BouncingScrollPhysics(),
    children: const [],
  );
  print(pv.physics is BouncingScrollPhysics);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn page_view_page_snapping() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final pv = PageView(
    pageSnapping: false,
    children: const [],
  );
  print(pv.pageSnapping);
}
"#
        ),
        vec!["false"]
    );
}

#[test]
fn page_view_on_page_changed() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final pv = PageView(
    onPageChanged: (int page) {},
    children: const [],
  );
  print(pv.onPageChanged != null);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn page_view_builder() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final pv = PageView.builder(
    itemCount: 5,
    itemBuilder: (BuildContext context, int index) => const SizedBox(),
  );
  print(pv is StatefulWidget);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn page_view_custom() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final pv = PageView.custom(
    childrenDelegate: SliverChildListDelegate(const []),
  );
  print(pv is StatefulWidget);
}
"#
        ),
        vec!["true"]
    );
}
