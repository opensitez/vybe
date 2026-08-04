// vybe-test: dart/flutter_ui_canvas_transformations/canvas_clip_path
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
  final path = Path()..addRect(Rect.fromLTRB(0, 0, 10, 10));
  canvas.clipPath(path);
  __p('clipped_path');
}

void main() {
  __vybeMain();
  __check('clipped_path');
}
