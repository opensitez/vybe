// vybe-test: dart/flutter_widgets_custom_paint/custom_painter_semantics
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
class MyPainter extends CustomPainter {
  @override
  void paint(Canvas canvas, Size size) {}
  @override
  bool shouldRepaint(covariant CustomPainter oldDelegate) => false;
  @override
  bool? hitTest(Offset position) => true;
}
void __vybeMain() {
  final p = MyPainter();
  __p(p.hitTest(Offset.zero));
}

void main() {
  __vybeMain();
  __check('true');
}
