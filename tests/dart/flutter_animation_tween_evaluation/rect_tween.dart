// vybe-test: dart/flutter_animation_tween_evaluation/rect_tween
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
  final tween = RectTween(
    begin: Rect.fromLTRB(0, 0, 10, 10),
    end: Rect.fromLTRB(10, 10, 30, 30),
  );
  final r = tween.lerp(0.5);
  __p('${r?.left}:${r?.right}');
}

void main() {
  __vybeMain();
  __check('5.0:20.0');
}
