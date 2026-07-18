use super::helpers::run_prints;

// ═══════════════════════════════════════════════════════════
// Flutter: widgets FutureBuilder
// ═══════════════════════════════════════════════════════════

#[test]
fn future_builder_creation() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final fb = FutureBuilder<String>(
    future: Future.value('Hello'),
    builder: (BuildContext context, AsyncSnapshot<String> snapshot) {
      return const SizedBox();
    },
  );
  print(fb is StatefulWidget);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn future_builder_initial_data() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final fb = FutureBuilder<int>(
    initialData: 42,
    future: Future.value(100),
    builder: (BuildContext context, AsyncSnapshot<int> snapshot) {
      return const SizedBox();
    },
  );
  print(fb.initialData);
}
"#
        ),
        vec!["42"]
    );
}

#[test]
fn future_builder_future_null() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final fb = FutureBuilder<bool>(
    future: null,
    builder: (BuildContext context, AsyncSnapshot<bool> snapshot) {
      return const SizedBox();
    },
  );
  print(fb.future == null);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn async_snapshot_nothing() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  const snapshot = AsyncSnapshot<String>.nothing();
  print('${snapshot.connectionState.name}:${snapshot.hasData}');
}
"#
        ),
        vec!["none:false"]
    );
}

#[test]
fn async_snapshot_waiting() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  const snapshot = AsyncSnapshot<int>.waiting();
  print('${snapshot.connectionState.name}:${snapshot.hasData}');
}
"#
        ),
        vec!["waiting:false"]
    );
}

#[test]
fn async_snapshot_with_data() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  const snapshot = AsyncSnapshot<String>.withData(ConnectionState.done, 'Result');
  print('${snapshot.connectionState.name}:${snapshot.data}');
}
"#
        ),
        vec!["done:Result"]
    );
}

#[test]
fn async_snapshot_with_error() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  const snapshot = AsyncSnapshot<double>.withError(ConnectionState.done, 'Error', StackTrace.empty);
  print('${snapshot.connectionState.name}:${snapshot.hasError}');
}
"#
        ),
        vec!["done:true"]
    );
}
