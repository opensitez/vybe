// vybe-test: dart/flutter_ui_color_math/color_tostring
// origin: languages/dart/tests/dart/test_flutter_ui_color_math.rs

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
  final c = Color(0xFF123456);
  __p(c.toString().contains('123456'));
}

void main() {
  __vybeMain();
  __check('true');
}
