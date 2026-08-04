// vybe-test: dart/flutter_widgets_positioned/positioned_right_bottom
// origin: languages/dart/tests/dart/test_flutter_widgets_positioned.rs

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
  final p = Positioned(right: 30.0, bottom: 40.0, child: const SizedBox());
  __p('${p.right}:${p.bottom}');
}

void main() {
  __vybeMain();
  __check('30.0:40.0');
}
