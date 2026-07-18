use super::helpers::run_prints;

// ═══════════════════════════════════════════════════════════
// Flutter: widgets SliverToBoxAdapter
// ═══════════════════════════════════════════════════════════

#[test]
fn sliver_to_box_adapter_creation() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  const stba = SliverToBoxAdapter(
    child: SizedBox(),
  );
  print(stba is SingleChildRenderObjectWidget);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn sliver_to_box_adapter_child() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  const stba = SliverToBoxAdapter(
    child: Text('Adapter'),
  );
  print(stba.child is Text);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn sliver_to_box_adapter_null_child() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  const stba = SliverToBoxAdapter();
  print(stba.child == null);
}
"#
        ),
        vec!["true"]
    );
}
