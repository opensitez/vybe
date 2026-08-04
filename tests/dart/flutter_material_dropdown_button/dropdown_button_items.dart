// vybe-test: dart/flutter_material_dropdown_button/dropdown_button_items
// origin: languages/dart/tests/dart/test_flutter_material_dropdown_button.rs

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
  final db = DropdownButton<String>(
    items: const [
      DropdownMenuItem(value: 'A', child: Text('A')),
      DropdownMenuItem(value: 'B', child: Text('B')),
    ],
    onChanged: (String? newValue) {},
  );
  __p(db.items?.length);
}

void main() {
  __vybeMain();
  __check('2');
}
