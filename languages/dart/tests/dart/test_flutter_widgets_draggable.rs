use super::helpers::run_prints;

// ═══════════════════════════════════════════════════════════
// Flutter: widgets Draggable
// ═══════════════════════════════════════════════════════════

#[test]
fn draggable_creation() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final d = Draggable<int>(
    data: 1,
    feedback: const SizedBox(),
    child: const SizedBox(),
  );
  print(d is StatefulWidget);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn draggable_data() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final d = Draggable<String>(
    data: 'Data',
    feedback: const SizedBox(),
    child: const SizedBox(),
  );
  print(d.data);
}
"#
        ),
        vec!["Data"]
    );
}

#[test]
fn draggable_feedback() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final d = Draggable<int>(
    feedback: const Text('Feedback'),
    child: const SizedBox(),
  );
  print(d.feedback is Text);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn draggable_child_when_dragging() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final d = Draggable<int>(
    childWhenDragging: const Placeholder(),
    feedback: const SizedBox(),
    child: const SizedBox(),
  );
  print(d.childWhenDragging is Placeholder);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn drag_target_creation() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final dt = DragTarget<int>(
    builder: (BuildContext context, List<int?> candidateData, List<dynamic> rejectedData) {
      return const SizedBox();
    },
  );
  print(dt is StatefulWidget);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn drag_target_on_will_accept() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final dt = DragTarget<int>(
    onWillAcceptWithDetails: (DragTargetDetails<int> details) => true,
    builder: (BuildContext context, List<int?> candidateData, List<dynamic> rejectedData) {
      return const SizedBox();
    },
  );
  print(dt.onWillAcceptWithDetails != null);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn drag_target_on_accept() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final dt = DragTarget<int>(
    onAcceptWithDetails: (DragTargetDetails<int> details) {},
    builder: (BuildContext context, List<int?> candidateData, List<dynamic> rejectedData) {
      return const SizedBox();
    },
  );
  print(dt.onAcceptWithDetails != null);
}
"#
        ),
        vec!["true"]
    );
}
