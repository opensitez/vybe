use super::helpers::run_prints;

// ═══════════════════════════════════════════════════════════
// Flutter: widgets StatefulBuilder
// ═══════════════════════════════════════════════════════════

#[test]
fn stateful_builder_creation() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final sb = StatefulBuilder(
    builder: (BuildContext context, StateSetter setState) {
      return const SizedBox();
    },
  );
  print(sb is StatefulWidget);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn stateful_builder_builder_function() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final sb = StatefulBuilder(
    builder: (BuildContext context, StateSetter setState) {
      return const Text('Stateful');
    },
  );
  print(sb.builder != null);
}
"#
        ),
        vec!["true"]
    );
}
