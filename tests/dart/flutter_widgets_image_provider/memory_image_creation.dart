// vybe-test: dart/flutter_widgets_image_provider/memory_image_creation
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
import 'dart:typed_data';
void __vybeMain() {
  final bytes = Uint8List.fromList([1, 2, 3]);
  final img = MemoryImage(bytes);
  __p(img.bytes.length);
}

void main() {
  __vybeMain();
  __check('3');
}
