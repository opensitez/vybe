// vybe-test: dart/flutter_widgets_align/animated_align_is_implicit_animation
// origin: languages/dart/tests/dart/test_flutter_widgets_align.rs

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
void __vybeMain() {
  final a = AnimatedAlign(
    alignment: Alignment.center,
    duration: const Duration(seconds: 1),
  );
  __p(a is ImplicitlyAnimatedWidget);
}

void main() {
  __vybeMain();
  __check('true');
}
