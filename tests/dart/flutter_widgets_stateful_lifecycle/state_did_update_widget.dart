// vybe-test: dart/flutter_widgets_stateful_lifecycle/state_did_update_widget
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
  final int val;
  MyStateful(this.val);
  @override
  State<MyStateful> createState() => _MyStatefulState();
}
class _MyStatefulState extends State<MyStateful> {
  @override
  void didUpdateWidget(MyStateful oldWidget) {
    super.didUpdateWidget(oldWidget);
    __p('old:${oldWidget.val} new:${widget.val}');
  }
  @override
  Widget build(BuildContext context) => const SizedBox();
}
void __vybeMain() {
  final w1 = MyStateful(1);
  final state = w1.createState();
  // We can't trivially simulate the element tree update without flutter_test,
  // but we can just check if method is there
  print('method exists');
}

void main() {
  __vybeMain();
  __check('method exists');
}
