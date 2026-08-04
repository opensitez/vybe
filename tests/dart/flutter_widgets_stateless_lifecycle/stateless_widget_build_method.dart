// vybe-test: dart/flutter_widgets_stateless_lifecycle/stateless_widget_build_method
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
class MyWidget extends StatelessWidget {
  @override
  Widget build(BuildContext context) {
    __p('building');
    return const SizedBox();
  }
}
void __vybeMain() {
  final w = MyWidget();
  // BuildContext mock is often just passing null in naive tests, though flutter will complain
  // We just test if method exists and is callable
  try {
    w.build(null as dynamic);
  } catch(e) {
    // some widgets assert context != null
    __p('failed');
  }
}

void main() {
  __vybeMain();
  __check('building');
}
