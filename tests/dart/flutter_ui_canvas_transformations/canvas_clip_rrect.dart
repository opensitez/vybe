// vybe-test: dart/flutter_ui_canvas_transformations/canvas_clip_rrect
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
  final rrect = RRect.fromRectAndRadius(Rect.fromLTRB(0, 0, 10, 10), Radius.circular(2.0));
  canvas.clipRRect(rrect);
  __p('clipped_rrect');
}

void main() {
  __vybeMain();
  __check('clipped_rrect');
}
