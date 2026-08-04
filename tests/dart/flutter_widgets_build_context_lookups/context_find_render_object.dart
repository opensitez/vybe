// vybe-test: dart/flutter_widgets_build_context_lookups/context_find_render_object
// origin: languages/dart/tests/dart/test_flutter_widgets_build_context_lookups.rs

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
    // In real app, this finds the render object
    __p(context.findRenderObject() == null);
    return const SizedBox();
  }
}
void __vybeMain() {
  final w = MyWidget();
  final e = w.createElement();
  // e.findRenderObject() will return null because not mounted
  __p(e.findRenderObject() == null);
}

void main() {
  __vybeMain();
  __check('true');
}
