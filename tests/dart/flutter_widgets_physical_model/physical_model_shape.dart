// vybe-test: dart/flutter_widgets_physical_model/physical_model_shape
// origin: languages/dart/tests/dart/test_flutter_widgets_physical_model.rs

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

import 'package:flutter/widgets.dart';
void __vybeMain() {
  final pm = PhysicalModel(
    color: const Color(0xFF000000),
    shape: BoxShape.circle,
  );
  __p(pm.shape == BoxShape.circle);
}

void main() {
  __vybeMain();
  __check('true');
}
