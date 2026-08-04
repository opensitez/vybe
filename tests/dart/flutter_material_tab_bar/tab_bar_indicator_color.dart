// vybe-test: dart/flutter_material_tab_bar/tab_bar_indicator_color
// origin: languages/dart/tests/dart/test_flutter_material_tab_bar.rs

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
  final t = TabBar(
    indicatorColor: const Color(0xFF00FF00),
    tabs: const [Tab(text: 'A')],
  );
  __p(t.indicatorColor?.value == 0xFF00FF00);
}

void main() {
  __vybeMain();
  __check('true');
}
