// vybe-test: dart/flutter_ui_color_math/color_compute_luminance
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
  final c = Color(0xFFFFFFFF);
  __p(c.computeLuminance() == 1.0);
  final b = Color(0xFF000000);
  __p(b.computeLuminance() == 0.0);
}

void main() {
  __vybeMain();
  __check('true\ntrue');
}
