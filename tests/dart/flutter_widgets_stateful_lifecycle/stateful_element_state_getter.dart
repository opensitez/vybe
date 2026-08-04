// vybe-test: dart/flutter_widgets_stateful_lifecycle/stateful_element_state_getter
// origin: languages/dart/tests/dart/test_flutter_widgets_stateful_lifecycle.rs

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
class MyStateful extends StatefulWidget {
  @override
  State<MyStateful> createState() => _MyStatefulState();
}
class _MyStatefulState extends State<MyStateful> {
  @override
  Widget build(BuildContext context) => const SizedBox();
}
void __vybeMain() {
  final w = MyStateful();
  final e = w.createElement();
  __p(e.state is _MyStatefulState);
}

void main() {
  __vybeMain();
  __check('true');
}
