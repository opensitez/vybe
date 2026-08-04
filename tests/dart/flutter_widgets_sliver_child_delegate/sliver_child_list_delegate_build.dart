// vybe-test: dart/flutter_widgets_sliver_child_delegate/sliver_child_list_delegate_build
// origin: languages/dart/tests/dart/test_flutter_widgets_sliver_child_delegate.rs

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
  final w1 = const SizedBox();
  final w2 = const Placeholder();
  final delegate = SliverChildListDelegate([w1, w2]);
  final e = const SizedBox().createElement();
  __p(delegate.build(e, 1) == w2);
}

void main() {
  __vybeMain();
  __check('true');
}
