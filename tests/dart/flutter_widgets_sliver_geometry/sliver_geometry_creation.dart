// vybe-test: dart/flutter_widgets_sliver_geometry/sliver_geometry_creation
// origin: languages/dart/tests/dart/test_flutter_widgets_sliver_geometry.rs

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
  final sg = SliverGeometry(scrollExtent: 100.0, paintExtent: 50.0);
  __p('${sg.scrollExtent}:${sg.paintExtent}');
}

void main() {
  __vybeMain();
  __check('100.0:50.0');
}
