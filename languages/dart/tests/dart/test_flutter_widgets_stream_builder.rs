use super::helpers::run_prints;

// ═══════════════════════════════════════════════════════════
// Flutter: widgets StreamBuilder
// ═══════════════════════════════════════════════════════════

#[test]
fn stream_builder_creation() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final sb = StreamBuilder<String>(
    stream: Stream.value('Hello'),
    builder: (BuildContext context, AsyncSnapshot<String> snapshot) {
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
fn stream_builder_initial_data() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final sb = StreamBuilder<int>(
    initialData: 0,
    stream: Stream.value(1),
    builder: (BuildContext context, AsyncSnapshot<int> snapshot) {
      return const SizedBox();
    },
  );
  print(sb.initialData);
}
"#
        ),
        vec!["0"]
    );
}

#[test]
fn stream_builder_stream_null() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final sb = StreamBuilder<double>(
    stream: null,
    builder: (BuildContext context, AsyncSnapshot<double> snapshot) {
      return const SizedBox();
    },
  );
  print(sb.stream == null);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn stream_builder_builder_function() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final sb = StreamBuilder<bool>(
    builder: (BuildContext context, AsyncSnapshot<bool> snapshot) {
      return const Text('Builder');
    },
  );
  print(sb.builder != null);
}
"#
        ),
        vec!["true"]
    );
}
