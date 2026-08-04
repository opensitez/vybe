// vybe-test: dart/flutter_ui_rect_rrect/rect_inflate
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
  final r = Rect.fromLTRB(10.0, 10.0, 20.0, 20.0);
  final inf = r.inflate(5.0);
  __p('${inf.left}:${inf.right}');
}

void main() {
  __vybeMain();
  __check('5.0:25.0');
}
