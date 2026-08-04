// vybe-test: dart/flutter_material_stepper/stepper_creation
// origin: languages/dart/tests/dart/test_flutter_material_stepper.rs

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
  final s = Stepper(steps: const []);
  __p(s is StatefulWidget);
}

void main() {
  __vybeMain();
  __check('true');
}
