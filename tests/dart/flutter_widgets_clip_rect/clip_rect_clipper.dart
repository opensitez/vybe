// vybe-test: dart/flutter_widgets_clip_rect/clip_rect_clipper
// origin: languages/dart/tests/dart/test_flutter_widgets_clip_rect.rs

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
class MyClipper extends CustomClipper<Rect> {
  @override
  Rect getClip(Size size) => Rect.fromLTWH(0, 0, 50, 50);
  @override
  bool shouldReclip(oldClipper) => false;
}
void __vybeMain() {
  final cr = ClipRect(clipper: MyClipper());
  __p(cr.clipper != null);
}

void main() {
  __vybeMain();
  __check('true');
}
