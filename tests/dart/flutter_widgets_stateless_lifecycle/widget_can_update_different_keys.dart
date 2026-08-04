// vybe-test: dart/flutter_widgets_stateless_lifecycle/widget_can_update_different_keys
// origin: languages/dart/tests/dart/test_flutter_widgets_stateless_lifecycle.rs

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
  final w1 = SizedBox(key: ValueKey(1));
  final w2 = SizedBox(key: ValueKey(2));
  __p(Widget.canUpdate(w1, w2));
}

void main() {
  __vybeMain();
  __check('false');
}
