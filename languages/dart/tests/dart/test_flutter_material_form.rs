use super::helpers::run_prints;

// ═══════════════════════════════════════════════════════════
// Flutter: material Form
// ═══════════════════════════════════════════════════════════

#[test]
fn form_creation() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/material.dart';
void main() {
  final f = Form(child: const SizedBox());
  print(f is StatefulWidget);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn form_child() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/material.dart';
void main() {
  final f = Form(child: const Placeholder());
  print(f.child is Placeholder);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn form_autovalidate_mode() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/material.dart';
void main() {
  final f = Form(
    autovalidateMode: AutovalidateMode.always,
    child: const SizedBox(),
  );
  print(f.autovalidateMode == AutovalidateMode.always);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn form_on_changed() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/material.dart';
void main() {
  final f = Form(
    onChanged: () {},
    child: const SizedBox(),
  );
  print(f.onChanged != null);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn form_can_pop() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/material.dart';
void main() {
  final f = Form(
    canPop: false,
    child: const SizedBox(),
  );
  print(f.canPop);
}
"#
        ),
        vec!["false"]
    );
}

#[test]
fn form_on_pop_invoked_with_result() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/material.dart';
void main() {
  final f = Form(
    onPopInvokedWithResult: (bool didPop, Object? result) {},
    child: const SizedBox(),
  );
  print(f.onPopInvokedWithResult != null);
}
"#
        ),
        vec!["true"]
    );
}
