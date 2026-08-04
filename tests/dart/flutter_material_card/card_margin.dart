// vybe-test: dart/flutter_material_card/card_margin
// origin: languages/dart/tests/dart/test_flutter_material_card.rs

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
  final c = Card(margin: const EdgeInsets.all(8.0));
  __p((c.margin as EdgeInsets).top);
}

void main() {
  __vybeMain();
  __check('8.0');
}
