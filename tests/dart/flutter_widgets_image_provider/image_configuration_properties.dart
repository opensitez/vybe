// vybe-test: dart/flutter_widgets_image_provider/image_configuration_properties
// origin: languages/dart/tests/dart/test_flutter_widgets_image_provider.rs

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
  final c = ImageConfiguration(size: Size(100, 100), devicePixelRatio: 2.0);
  __p('${c.size!.width}:${c.devicePixelRatio}');
}

void main() {
  __vybeMain();
  __check('100.0:2.0');
}
