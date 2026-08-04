// vybe-test: dart/flutter_widgets_center/center_width_factor
// origin: languages/dart/tests/dart/test_flutter_widgets_center.rs

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
  final c = Center(widthFactor: 3.0, child: const SizedBox());
  __p(c.widthFactor);
}

void main() {
  __vybeMain();
  __check('3.0');
}
