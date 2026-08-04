// vybe-test: dart/flutter_ui_canvas_transformations/canvas_draw_line
// origin: languages/dart/tests/dart/test_flutter_ui_canvas_transformations.rs

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
  final recorder = PictureRecorder();
  final canvas = Canvas(recorder);
  final paint = Paint();
  canvas.drawLine(Offset(0, 0), Offset(10, 10), paint);
  __p('drew_line');
}

void main() {
  __vybeMain();
  __check('drew_line');
}
