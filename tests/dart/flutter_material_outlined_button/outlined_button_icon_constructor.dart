// vybe-test: dart/flutter_material_outlined_button/outlined_button_icon_constructor
// origin: languages/dart/tests/dart/test_flutter_material_outlined_button.rs

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
  final ob = OutlinedButton.icon(
    onPressed: () {},
    icon: const Icon(Icons.download),
    label: const Text('Download'),
  );
  __p(ob is StatefulWidget);
}

void main() {
  __vybeMain();
  __check('true');
}
