// vybe-test: dart/flutter_widgets_fitted_box/apply_box_fit
// origin: languages/dart/tests/dart/test_flutter_widgets_fitted_box.rs

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

import 'package:flutter/painting.dart';
import 'dart:ui';
void __vybeMain() {
  final fs = applyBoxFit(BoxFit.contain, Size(100, 100), Size(50, 50));
  __p('${fs.source.width}:${fs.destination.width}');
}

void main() {
  __vybeMain();
  __check('100.0:50.0');
}
