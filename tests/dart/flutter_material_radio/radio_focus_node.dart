// vybe-test: dart/flutter_material_radio/radio_focus_node
// origin: languages/dart/tests/dart/test_flutter_material_radio.rs

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
  final fn = FocusNode();
  final r = Radio<int>(
    value: 1,
    groupValue: 1,
    focusNode: fn,
    onChanged: (int? newValue) {},
  );
  __p(r.focusNode == fn);
}

void main() {
  __vybeMain();
  __check('true');
}
