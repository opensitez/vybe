// vybe-test: dart/flutter_material_bottom_sheet/bottom_sheet_elevation
// origin: languages/dart/tests/dart/test_flutter_material_bottom_sheet.rs

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
  final bs = BottomSheet(
    elevation: 8.0,
    onClosing: () {},
    builder: (context) => const SizedBox(),
  );
  __p(bs.elevation);
}

void main() {
  __vybeMain();
  __check('8.0');
}
