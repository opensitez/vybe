// vybe-test: dart/flutter_ui_vertices_blend_mode/paint_blend_mode_assignment
// origin: languages/dart/tests/dart/test_flutter_ui_vertices_blend_mode.rs

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
  final paint = Paint()..blendMode = BlendMode.dstIn;
  __p(paint.blendMode.name);
}

void main() {
  __vybeMain();
  __check('dstIn');
}
