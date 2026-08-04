// vybe-test: dart/flutter_material_slider/slider_min_max
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
    value: 50.0,
    min: 10.0,
    max: 100.0,
    onChanged: (double newValue) {},
  );
  __p('${s.min}:${s.max}');
}

void main() {
  __vybeMain();
  __check('10.0:100.0');
}
