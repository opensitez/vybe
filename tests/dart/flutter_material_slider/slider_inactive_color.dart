// vybe-test: dart/flutter_material_slider/slider_inactive_color
// origin: languages/dart/tests/dart/test_flutter_material_slider.rs

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

import 'package:flutter/material.dart';
void __vybeMain() {
  final s = Slider(
    value: 0.5,
    inactiveColor: const Color(0xFF112233),
    onChanged: (double newValue) {},
  );
  __p(s.inactiveColor?.value == 0xFF112233);
}

void main() {
  __vybeMain();
  __check('true');
}
