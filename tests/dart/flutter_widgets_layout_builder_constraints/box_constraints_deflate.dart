// vybe-test: dart/flutter_widgets_layout_builder_constraints/box_constraints_deflate
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
import 'package:flutter/painting.dart';
void __vybeMain() {
  final c = BoxConstraints(minWidth: 100, maxWidth: 200, minHeight: 100, maxHeight: 200);
  final insets = EdgeInsets.all(10);
  final deflated = c.deflate(insets);
  __p('${deflated.minWidth}:${deflated.maxWidth}');
}

void main() {
  __vybeMain();
  __check('80.0:180.0');
}
