// vybe-test: dart/flutter_widgets_positioned/positioned_from_rect
// origin: languages/dart/tests/dart/test_flutter_widgets_positioned.rs

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
void __vybeMain() {
  final rect = Rect.fromLTRB(10.0, 20.0, 30.0, 40.0);
  final p = Positioned.fromRect(rect: rect, child: const SizedBox());
  __p('${p.left}:${p.top}:${p.width}:${p.height}');
}

void main() {
  __vybeMain();
  __check('10.0:20.0:20.0:20.0');
}
