// vybe-test: dart/flutter_widgets_image_widget/image_color_blend_mode
// origin: languages/dart/tests/dart/test_flutter_widgets_image_widget.rs

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
  final img = Image.memory(Uint8List(0), color: const Color(0xFFFF0000), colorBlendMode: BlendMode.srcOver);
  __p(img.colorBlendMode == BlendMode.srcOver);
}

void main() {
  __vybeMain();
  __check('true');
}
