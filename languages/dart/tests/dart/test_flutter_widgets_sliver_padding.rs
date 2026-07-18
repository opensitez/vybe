use super::helpers::run_prints;

// ═══════════════════════════════════════════════════════════
// Flutter: widgets SliverPadding
// ═══════════════════════════════════════════════════════════

#[test]
fn sliver_padding_creation() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  const sp = SliverPadding(
    padding: EdgeInsets.all(8.0),
  );
  print(sp is SingleChildRenderObjectWidget);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn sliver_padding_padding() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  const sp = SliverPadding(
    padding: EdgeInsets.all(12.0),
  );
  print((sp.padding as EdgeInsets).top);
}
"#
        ),
        vec!["12.0"]
    );
}

#[test]
fn sliver_padding_sliver() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final sp = SliverPadding(
    padding: const EdgeInsets.all(8.0),
    sliver: SliverList(delegate: SliverChildListDelegate(const [])),
  );
  print(sp.sliver is SliverList);
}
"#
        ),
        vec!["true"]
    );
}
