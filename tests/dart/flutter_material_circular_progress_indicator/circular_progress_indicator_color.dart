// vybe-test: dart/flutter_material_circular_progress_indicator/circular_progress_indicator_color
// origin: languages/dart/tests/dart/test_flutter_material_circular_progress_indicator.rs

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
  const cpi = CircularProgressIndicator(color: Color(0xFFFF0000));
  __p(cpi.color?.value == 0xFFFF0000);
}

void main() {
  __vybeMain();
  __check('true');
}
