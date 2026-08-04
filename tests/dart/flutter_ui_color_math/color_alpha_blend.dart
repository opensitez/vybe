// vybe-test: dart/flutter_ui_color_math/color_alpha_blend
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
  final fg = Color(0x80FF0000); // semi-transparent red
  final bg = Color(0xFF00FF00); // solid green
  final blended = Color.alphaBlend(fg, bg);
  __p(blended.alpha == 255);
}

void main() {
  __vybeMain();
  __check('true');
}
