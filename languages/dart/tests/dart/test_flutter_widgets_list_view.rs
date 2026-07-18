use super::helpers::run_prints;

// ═══════════════════════════════════════════════════════════
// Flutter: widgets ListView
// ═══════════════════════════════════════════════════════════

#[test]
fn list_view_creation() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final lv = ListView(children: const []);
  print(lv is BoxScrollView);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn list_view_scroll_direction() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final lv = ListView(scrollDirection: Axis.horizontal, children: const []);
  print(lv.scrollDirection == Axis.horizontal);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn list_view_reverse() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final lv = ListView(reverse: true, children: const []);
  print(lv.reverse);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn list_view_builder() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final lv = ListView.builder(
    itemCount: 10,
    itemBuilder: (BuildContext context, int index) => const SizedBox(),
  );
  print(lv is BoxScrollView);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn list_view_separated() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final lv = ListView.separated(
    itemCount: 5,
    itemBuilder: (BuildContext context, int index) => const SizedBox(),
    separatorBuilder: (BuildContext context, int index) => const SizedBox(),
  );
  print(lv is BoxScrollView);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn list_view_custom() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final lv = ListView.custom(
    childrenDelegate: SliverChildListDelegate(const []),
  );
  print(lv is BoxScrollView);
}
"#
        ),
        vec!["true"]
    );
}
