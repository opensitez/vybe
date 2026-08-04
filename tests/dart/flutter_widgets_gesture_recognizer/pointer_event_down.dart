// vybe-test: dart/flutter_widgets_gesture_recognizer/pointer_event_down
// origin: languages/dart/tests/dart/test_flutter_widgets_gesture_recognizer.rs

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

import 'package:flutter/gestures.dart';
void __vybeMain() {
  final pe = PointerDownEvent(pointer: 1);
  __p(pe.pointer);
}

void main() {
  __vybeMain();
  __check('1');
}
