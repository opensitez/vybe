// vybe-test: dart/flutter_widgets_physical_shape/physical_shape_clip_behavior
// origin: languages/dart/tests/dart/test_flutter_widgets_physical_shape.rs

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
import 'dart:ui';
class MyPathClipper extends CustomClipper<Path> {
  @override
  Path getClip(Size size) => Path()..addRect(Rect.fromLTWH(0,0,50,50));
  @override
  bool shouldReclip(oldClipper) => false;
}
void __vybeMain() {
  final ps = PhysicalShape(
    clipper: MyPathClipper(),
    color: const Color(0xFF000000),
    clipBehavior: Clip.hardEdge,
  );
  __p(ps.clipBehavior == Clip.hardEdge);
}

void main() {
  __vybeMain();
  __check('true');
}
