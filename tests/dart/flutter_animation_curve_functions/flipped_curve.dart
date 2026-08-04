// vybe-test: dart/flutter_animation_curve_functions/flipped_curve
// origin: languages/dart/tests/dart/test_flutter_animation_curve_functions.rs

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

import 'package:flutter/animation.dart';
void __vybeMain() {
  final flipped = FlippedCurve(Curves.easeIn);
  // Same as easeOut shape but exact values depend on math
  __p(flipped.transform(0.5) > 0.5);
}

void main() {
  __vybeMain();
  __check('true');
}
