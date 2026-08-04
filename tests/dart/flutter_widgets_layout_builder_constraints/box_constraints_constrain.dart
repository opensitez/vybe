// vybe-test: dart/flutter_widgets_layout_builder_constraints/box_constraints_constrain
// origin: languages/dart/tests/dart/test_flutter_widgets_layout_builder_constraints.rs

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

import 'package:flutter/rendering.dart';
import 'dart:ui';
void __vybeMain() {
  final c = BoxConstraints(minWidth: 50, maxWidth: 100, minHeight: 50, maxHeight: 100);
  final s1 = c.constrain(Size(10, 10)); // too small
  final s2 = c.constrain(Size(200, 200)); // too large
  __p('${s1.width}:${s1.height} ${s2.width}:${s2.height}');
}

void main() {
  __vybeMain();
  __check('50.0:50.0 100.0:100.0');
}
