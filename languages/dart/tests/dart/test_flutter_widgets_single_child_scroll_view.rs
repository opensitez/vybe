use super::helpers::run_prints;

// ═══════════════════════════════════════════════════════════
// Flutter: widgets SingleChildScrollView
// ═══════════════════════════════════════════════════════════

#[test]
fn single_child_scroll_view_creation() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  const scsv = SingleChildScrollView(child: SizedBox());
  print(scsv is StatelessWidget);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn single_child_scroll_view_scroll_direction() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  const scsv = SingleChildScrollView(
    scrollDirection: Axis.horizontal,
    child: SizedBox(),
  );
  print(scsv.scrollDirection == Axis.horizontal);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn single_child_scroll_view_reverse() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  const scsv = SingleChildScrollView(
    reverse: true,
    child: SizedBox(),
  );
  print(scsv.reverse);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn single_child_scroll_view_padding() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  const scsv = SingleChildScrollView(
    padding: EdgeInsets.all(10.0),
    child: SizedBox(),
  );
  print((scsv.padding as EdgeInsets).top);
}
"#
        ),
        vec!["10.0"]
    );
}

#[test]
fn single_child_scroll_view_primary() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  const scsv = SingleChildScrollView(
    primary: true,
    child: SizedBox(),
  );
  print(scsv.primary);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn single_child_scroll_view_physics() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  const scsv = SingleChildScrollView(
    physics: BouncingScrollPhysics(),
    child: SizedBox(),
  );
  print(scsv.physics is BouncingScrollPhysics);
}
"#
        ),
        vec!["true"]
    );
}
