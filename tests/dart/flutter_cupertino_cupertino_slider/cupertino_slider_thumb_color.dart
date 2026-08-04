// vybe-test: dart/flutter_cupertino_cupertino_slider/cupertino_slider_thumb_color
// origin: languages/dart/tests/dart/test_flutter_cupertino_cupertino_slider.rs

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

import 'package:flutter/cupertino.dart';
void __vybeMain() {
  final cs = CupertinoSlider(
    value: 0.5,
    thumbColor: const Color(0xFF123456),
    onChanged: (double newValue) {},
  );
  __p(cs.thumbColor.value == 0xFF123456);
}

void main() {
  __vybeMain();
  __check('true');
}
