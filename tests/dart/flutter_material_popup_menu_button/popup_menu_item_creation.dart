// vybe-test: dart/flutter_material_popup_menu_button/popup_menu_item_creation
// origin: languages/dart/tests/dart/test_flutter_material_popup_menu_button.rs

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
  final pmi = PopupMenuItem<double>(
    value: 3.14,
    child: const Text('Pi'),
  );
  __p('${pmi.value}:${pmi.child is Text}');
}

void main() {
  __vybeMain();
  __check('3.14:true');
}
