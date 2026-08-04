// vybe-test: dart/flutter_widgets_fractional_translation/fractional_translation_transform_hit_tests
// origin: languages/dart/tests/dart/test_flutter_widgets_fractional_translation.rs

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
  final ft = FractionalTranslation(
    translation: const Offset(1.0, 1.0),
    transformHitTests: false,
    child: const SizedBox(),
  );
  __p(ft.transformHitTests);
}

void main() {
  __vybeMain();
  __check('false');
}
