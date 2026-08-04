// vybe-test: dart/flutter_ui_rect_rrect/rect_overlaps
// origin: languages/dart/tests/dart/test_flutter_ui_rect_rrect.rs

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
  final r1 = Rect.fromLTRB(0.0, 0.0, 10.0, 10.0);
  final r2 = Rect.fromLTRB(10.0, 10.0, 20.0, 20.0); // Touches edge
  // Overlaps strictly greater than
  __p(r1.overlaps(r2));
}

void main() {
  __vybeMain();
  __check('false');
}
