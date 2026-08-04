// vybe-test: dart/flutter_ui_rect_rrect/rect_contains
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
  final r = Rect.fromLTRB(0.0, 0.0, 10.0, 10.0);
  __p(r.contains(Offset(5.0, 5.0)));
  __p(r.contains(Offset(15.0, 5.0)));
}

void main() {
  __vybeMain();
  __check('true\nfalse');
}
