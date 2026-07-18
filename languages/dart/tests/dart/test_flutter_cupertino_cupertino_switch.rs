use super::helpers::run_prints;

// ═══════════════════════════════════════════════════════════
// Flutter: cupertino CupertinoSwitch
// ═══════════════════════════════════════════════════════════

#[test]
fn cupertino_switch_creation() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/cupertino.dart';
void main() {
  final cs = CupertinoSwitch(
    value: true,
    onChanged: (bool newValue) {},
  );
  print(cs is StatefulWidget);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn cupertino_switch_value() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/cupertino.dart';
void main() {
  final cs = CupertinoSwitch(
    value: false,
    onChanged: (bool newValue) {},
  );
  print(cs.value);
}
"#
        ),
        vec!["false"]
    );
}

#[test]
fn cupertino_switch_active_color() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/cupertino.dart';
void main() {
  final cs = CupertinoSwitch(
    value: true,
    activeColor: const Color(0xFF00FF00),
    onChanged: (bool newValue) {},
  );
  print(cs.activeColor?.value == 0xFF00FF00);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn cupertino_switch_track_color() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/cupertino.dart';
void main() {
  final cs = CupertinoSwitch(
    value: false,
    trackColor: const Color(0xFF112233),
    onChanged: (bool newValue) {},
  );
  print(cs.trackColor?.value == 0xFF112233);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn cupertino_switch_thumb_color() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/cupertino.dart';
void main() {
  final cs = CupertinoSwitch(
    value: true,
    thumbColor: const Color(0xFF123456),
    onChanged: (bool newValue) {},
  );
  print(cs.thumbColor?.value == 0xFF123456);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn cupertino_switch_on_changed() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/cupertino.dart';
void main() {
  final cs = CupertinoSwitch(
    value: true,
    onChanged: null,
  );
  print(cs.onChanged == null);
}
"#
        ),
        vec!["true"]
    );
}
