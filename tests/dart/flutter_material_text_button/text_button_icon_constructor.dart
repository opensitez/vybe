// vybe-test: dart/flutter_material_text_button/text_button_icon_constructor
// origin: languages/dart/tests/dart/test_flutter_material_text_button.rs

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
  final tb = TextButton.icon(
    onPressed: () {},
    icon: const Icon(Icons.share),
    label: const Text('Share'),
  );
  __p(tb is StatefulWidget);
}

void main() {
  __vybeMain();
  __check('true');
}
