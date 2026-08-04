// vybe-test: dart/flutter_animation_tween_evaluation/tween_transform
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
void __vybeMain() {
  final tween = Tween<double>(begin: 50.0, end: 100.0);
  __p(tween.transform(0.5));
}

void main() {
  __vybeMain();
  __check('75.0');
}
