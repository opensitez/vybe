// vybe-test: dart/flutter_widgets_layout_builder_constraints/box_constraints_tightFor
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
void __vybeMain() {
  final c = BoxConstraints.tightFor(width: 50);
  // height is 0 to infinity
  __p('${c.minWidth}:${c.maxWidth}:${c.minHeight}:${c.maxHeight == double.infinity}');
}

void main() {
  __vybeMain();
  __check('50.0:50.0:0.0:true');
}
