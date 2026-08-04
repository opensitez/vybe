// vybe-test: dart/flutter_widgets_clip_rrect/clip_rrect_clipper
// origin: languages/dart/tests/dart/test_flutter_widgets_clip_rrect.rs

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
class MyClipper extends CustomClipper<RRect> {
  @override
  RRect getClip(Size size) => RRect.fromLTRBR(0, 0, 50, 50, Radius.circular(5));
  @override
  bool shouldReclip(oldClipper) => false;
}
void __vybeMain() {
  final cr = ClipRRect(clipper: MyClipper());
  __p(cr.clipper != null);
}

void main() {
  __vybeMain();
  __check('true');
}
