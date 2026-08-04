// vybe-test: dart/flutter_animation_tween_evaluation/color_tween
// origin: languages/dart/tests/dart/test_flutter_animation_tween_evaluation.rs

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

import 'package:flutter/animation.dart';
import 'dart:ui';
void __vybeMain() {
  final tween = ColorTween(begin: Color(0xFF000000), end: Color(0xFFFFFFFF));
  final c = tween.lerp(0.5);
  // Color lerping is per-channel, so 0x7F or 0x80 usually
  __p(c?.alpha == 255);
  __p(c?.red == 127 || c?.red == 128);
}

void main() {
  __vybeMain();
  __check('true\ntrue');
}
