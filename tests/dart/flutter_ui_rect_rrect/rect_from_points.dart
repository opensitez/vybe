// vybe-test: dart/flutter_ui_rect_rrect/rect_from_points
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
  final p1 = Offset(20.0, 30.0);
  final p2 = Offset(10.0, 10.0); // Out of order coordinates
  final r = Rect.fromPoints(p1, p2);
  __p('${r.left}:${r.right}:${r.top}:${r.bottom}');
}

void main() {
  __vybeMain();
  __check('10.0:20.0:10.0:30.0');
}
