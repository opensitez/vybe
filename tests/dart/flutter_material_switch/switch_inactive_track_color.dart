// vybe-test: dart/flutter_material_switch/switch_inactive_track_color
// origin: languages/dart/tests/dart/test_flutter_material_switch.rs

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
  final s = Switch(
    value: false,
    inactiveTrackColor: const Color(0xFF778899),
    onChanged: (bool newValue) {},
  );
  __p(s.inactiveTrackColor?.value == 0xFF778899);
}

void main() {
  __vybeMain();
  __check('true');
}
