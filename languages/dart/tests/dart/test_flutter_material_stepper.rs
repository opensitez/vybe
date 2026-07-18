use super::helpers::run_prints;

// ═══════════════════════════════════════════════════════════
// Flutter: material Stepper
// ═══════════════════════════════════════════════════════════

#[test]
fn stepper_creation() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/material.dart';
void main() {
  final s = Stepper(steps: const []);
  print(s is StatefulWidget);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn stepper_steps() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/material.dart';
void main() {
  final s = Stepper(
    steps: const [
      Step(title: Text('1'), content: SizedBox()),
      Step(title: Text('2'), content: SizedBox()),
    ],
  );
  print(s.steps.length);
}
"#
        ),
        vec!["2"]
    );
}

#[test]
fn stepper_current_step() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/material.dart';
void main() {
  final s = Stepper(
    currentStep: 1,
    steps: const [
      Step(title: Text('1'), content: SizedBox()),
      Step(title: Text('2'), content: SizedBox()),
    ],
  );
  print(s.currentStep);
}
"#
        ),
        vec!["1"]
    );
}

#[test]
fn stepper_type() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/material.dart';
void main() {
  final s = Stepper(
    type: StepperType.horizontal,
    steps: const [],
  );
  print(s.type == StepperType.horizontal);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn step_creation() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/material.dart';
void main() {
  final step = Step(
    title: const Text('A'),
    subtitle: const Text('B'),
    content: const SizedBox(),
    isActive: true,
    state: StepState.complete,
  );
  print('${step.isActive}:${step.state.name}');
}
"#
        ),
        vec!["true:complete"]
    );
}
