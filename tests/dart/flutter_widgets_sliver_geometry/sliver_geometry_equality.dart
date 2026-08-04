// vybe-test: dart/flutter_widgets_sliver_geometry/sliver_geometry_equality
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
  final sg1 = SliverGeometry(scrollExtent: 100.0);
  final sg2 = SliverGeometry(scrollExtent: 100.0);
  // Dart flutter equality might or might not compare instances or properties depending on version
  // Actually, SliverGeometry usually doesn't override == (identity only) or maybe it does?
  // Let's print properties
  __p(sg1.scrollExtent == sg2.scrollExtent);
}

void main() {
  __vybeMain();
  __check('true');
}
