// vybe-test: dart/flutter_ui_paint_shader/paint_shader_sweep_gradient
// origin: languages/dart/tests/dart/test_flutter_ui_paint_shader.rs

final StringBuffer __vybeOut = StringBuffer();

void __p(Object? o) {
  __vybeOut.writeln(o);
}

void __check(String want) {
  var got = __vybeOut.toString();
  // `writeln` on the final print contributes a trailing newline that the
  // expected line vector never carried.
  if (got.endsWith('\n')) {
    got = got.substring(0, got.length - 1);
  }
  if (got != want) {
    print('FAIL: want [$want] got [$got]');
    throw Exception('assertion failed');
  }
}

import 'dart:ui';
void __vybeMain() {
  final paint = Paint();
  paint.shader = Gradient.sweep(
    Offset(5, 5),
    [Color(0xFF000000), Color(0xFFFFFFFF)],
  );
  __p(paint.shader != null);
}

void main() {
  __vybeMain();
  __check('true');
}
