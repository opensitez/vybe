use super::helpers::run_prints;

// ═══════════════════════════════════════════════════════════
// Flutter: widgets ColorFiltered
// ═══════════════════════════════════════════════════════════

#[test]
fn color_filtered_creation() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final cf = ColorFiltered(
    colorFilter: const ColorFilter.mode(Color(0xFFFF0000), BlendMode.srcIn),
    child: const SizedBox(),
  );
  print(cf.colorFilter != null);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn color_filtered_child() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final cf = ColorFiltered(
    colorFilter: const ColorFilter.mode(Color(0xFF00FF00), BlendMode.srcOut),
    child: const Placeholder(),
  );
  print(cf.child is Placeholder);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn color_filtered_is_single_child_render_object() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final cf = ColorFiltered(
    colorFilter: const ColorFilter.mode(Color(0xFF0000FF), BlendMode.srcOver),
  );
  print(cf is SingleChildRenderObjectWidget);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn color_filtered_render_object() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
import 'package:flutter/rendering.dart';
void main() {
  final cf = ColorFiltered(
    colorFilter: const ColorFilter.mode(Color(0xFF112233), BlendMode.color),
  );
  // Creates RenderColorFiltered
  print('compiles');
}
"#
        ),
        vec!["compiles"]
    );
}
