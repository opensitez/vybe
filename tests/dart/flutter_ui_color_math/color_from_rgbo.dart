// vybe-test: dart/flutter_ui_color_math/color_from_rgbo
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
  // opacity is 0.0 to 1.0
  final c = Color.fromRGBO(255, 0, 0, 0.5);
  // 0.5 * 255 = 127 = 0x7F
  __p(c.alpha == 127 || c.alpha == 128); // Precision might make it 127 or 128 depending on rounding
  __p('${c.red}:${c.green}:${c.blue}');
}

void main() {
  __vybeMain();
  __check('true\n255:0:0');
}
