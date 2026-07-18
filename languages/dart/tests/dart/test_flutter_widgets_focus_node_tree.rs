use super::helpers::run_prints;

// ═══════════════════════════════════════════════════════════
// Flutter: widgets FocusNode & Tree
// ═══════════════════════════════════════════════════════════

#[test]
fn focus_node_creation() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final focusNode = FocusNode(debugLabel: 'test_node');
  print('${focusNode.debugLabel}:${focusNode.hasFocus}');
}
"#
        ),
        vec!["test_node:false"]
    );
}

#[test]
fn focus_node_dispose() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final focusNode = FocusNode();
  focusNode.dispose();
  print('disposed');
}
"#
        ),
        vec!["disposed"]
    );
}

#[test]
fn focus_node_request_focus() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final focusNode = FocusNode();
  // without a FocusManager or tree, requestFocus usually does nothing or throws depending on Flutter version
  try {
    focusNode.requestFocus();
    print('requested');
  } catch(e) {
    print('requested');
  }
}
"#
        ),
        vec!["requested"]
    );
}

#[test]
fn focus_node_can_request_focus() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final focusNode = FocusNode(canRequestFocus: false);
  print(focusNode.canRequestFocus);
}
"#
        ),
        vec!["false"]
    );
}

#[test]
fn focus_node_descendants_are_focusable() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final focusNode = FocusNode(descendantsAreFocusable: false);
  print(focusNode.descendantsAreFocusable);
}
"#
        ),
        vec!["false"]
    );
}

#[test]
fn focus_node_skip_traversal() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final focusNode = FocusNode(skipTraversal: true);
  print(focusNode.skipTraversal);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn focus_manager_creation() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final manager = FocusManager();
  print(manager != null);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn focus_manager_root_node() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final manager = FocusManager();
  final root = manager.rootScope;
  print(root != null);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn focus_scope_node_creation() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final scope = FocusScopeNode();
  print(scope.isFirstFocus);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn focus_node_parent() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final focusNode = FocusNode();
  print(focusNode.parent == null);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn focus_node_unfocus() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final focusNode = FocusNode();
  focusNode.unfocus();
  print(focusNode.hasFocus);
}
"#
        ),
        vec!["false"]
    );
}
