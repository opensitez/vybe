// vybe-test: dart/flutter_animation_tween_evaluation/reverse_tween
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
  final tween = ReverseTween(Tween<double>(begin: 0.0, end: 10.0));
  // reverse tween uses (1.0 - t) for the parent
  __p(tween.lerp(0.25)); // parent sees 0.75, so 7.5
}

void main() {
  __vybeMain();
  __check('7.5');
}
