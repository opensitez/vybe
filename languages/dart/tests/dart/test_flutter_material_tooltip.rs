use super::helpers::run_prints;

// ═══════════════════════════════════════════════════════════
// Flutter: material Tooltip
// ═══════════════════════════════════════════════════════════

#[test]
fn tooltip_creation() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/material.dart';
void main() {
  const t = Tooltip(
    message: 'Hello',
    child: SizedBox(),
  );
  print(t is StatefulWidget);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn tooltip_message() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/material.dart';
void main() {
  const t = Tooltip(
    message: 'Test message',
    child: SizedBox(),
  );
  print(t.message);
}
"#
        ),
        vec!["Test message"]
    );
}

#[test]
fn tooltip_rich_message() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/material.dart';
void main() {
  const t = Tooltip(
    richMessage: TextSpan(text: 'Rich'),
    child: SizedBox(),
  );
  print((t.richMessage as TextSpan).text);
}
"#
        ),
        vec!["Rich"]
    );
}

#[test]
fn tooltip_height() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/material.dart';
void main() {
  const t = Tooltip(
    message: 'Test',
    height: 40.0,
    child: SizedBox(),
  );
  print(t.height);
}
"#
        ),
        vec!["40.0"]
    );
}

#[test]
fn tooltip_padding() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/material.dart';
void main() {
  const t = Tooltip(
    message: 'Test',
    padding: EdgeInsets.all(8.0),
    child: SizedBox(),
  );
  print((t.padding as EdgeInsets).top);
}
"#
        ),
        vec!["8.0"]
    );
}

#[test]
fn tooltip_wait_duration() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/material.dart';
void main() {
  const t = Tooltip(
    message: 'Test',
    waitDuration: Duration(seconds: 1),
    child: SizedBox(),
  );
  print(t.waitDuration?.inSeconds);
}
"#
        ),
        vec!["1"]
    );
}

#[test]
fn tooltip_show_duration() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/material.dart';
void main() {
  const t = Tooltip(
    message: 'Test',
    showDuration: Duration(seconds: 2),
    child: SizedBox(),
  );
  print(t.showDuration?.inSeconds);
}
"#
        ),
        vec!["2"]
    );
}
