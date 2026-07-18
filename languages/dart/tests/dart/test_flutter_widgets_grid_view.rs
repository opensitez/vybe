use super::helpers::run_prints;

// ═══════════════════════════════════════════════════════════
// Flutter: widgets GridView
// ═══════════════════════════════════════════════════════════

#[test]
fn grid_view_creation() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final gv = GridView(
    gridDelegate: const SliverGridDelegateWithFixedCrossAxisCount(crossAxisCount: 2),
    children: const [],
  );
  print(gv is BoxScrollView);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn grid_view_count() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final gv = GridView.count(
    crossAxisCount: 3,
    children: const [],
  );
  print(gv is BoxScrollView);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn grid_view_extent() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final gv = GridView.extent(
    maxCrossAxisExtent: 150.0,
    children: const [],
  );
  print(gv is BoxScrollView);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn grid_view_builder() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final gv = GridView.builder(
    gridDelegate: const SliverGridDelegateWithFixedCrossAxisCount(crossAxisCount: 2),
    itemBuilder: (BuildContext context, int index) => const SizedBox(),
  );
  print(gv is BoxScrollView);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn grid_view_custom() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final gv = GridView.custom(
    gridDelegate: const SliverGridDelegateWithFixedCrossAxisCount(crossAxisCount: 2),
    childrenDelegate: SliverChildListDelegate(const []),
  );
  print(gv is BoxScrollView);
}
"#
        ),
        vec!["true"]
    );
}
