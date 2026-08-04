// vybe-test: dart/flutter_widgets_custom_paint/custom_paint_creation
// origin: languages/dart/tests/dart/test_flutter_widgets_custom_paint.rs

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
import 'package:flutter/rendering.dart';
class MyPainter extends CustomPainter {
  @override
  void paint(Canvas canvas, Size size) {}
  @override
  bool shouldRepaint(covariant CustomPainter oldDelegate) => false;
}
void __vybeMain() {
  final cp = CustomPaint(painter: MyPainter());
  __p(cp.painter != null);
}

void main() {
  __vybeMain();
  __check('true');
}
