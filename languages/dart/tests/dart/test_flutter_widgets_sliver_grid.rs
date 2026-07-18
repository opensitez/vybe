use super::helpers::run_prints;

// ═══════════════════════════════════════════════════════════
// Flutter: widgets SliverGrid
// ═══════════════════════════════════════════════════════════

#[test]
fn sliver_grid_creation() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final sg = SliverGrid(
    delegate: SliverChildListDelegate(const []),
    gridDelegate: const SliverGridDelegateWithFixedCrossAxisCount(crossAxisCount: 2),
  );
  print(sg is StatelessWidget);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn sliver_grid_count() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final sg = SliverGrid.count(
    crossAxisCount: 3,
    children: const [],
  );
  print(sg is StatelessWidget);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn sliver_grid_extent() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final sg = SliverGrid.extent(
    maxCrossAxisExtent: 100.0,
    children: const [],
  );
  print(sg is StatelessWidget);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn sliver_grid_builder() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final sg = SliverGrid.builder(
    gridDelegate: const SliverGridDelegateWithFixedCrossAxisCount(crossAxisCount: 2),
    itemCount: 10,
    itemBuilder: (BuildContext context, int index) => const SizedBox(),
  );
  print(sg is StatelessWidget);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn sliver_grid_delegate() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final sg = SliverGrid(
    delegate: SliverChildListDelegate(const []),
    gridDelegate: const SliverGridDelegateWithFixedCrossAxisCount(crossAxisCount: 2),
  );
  print(sg.gridDelegate is SliverGridDelegate);
}
"#
        ),
        vec!["true"]
    );
}
