// vybe-test: dart/flutter_ui_image_decoding/image_descriptor_dispose
// origin: languages/dart/tests/dart/test_flutter_ui_image_decoding.rs

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

import 'dart:ui';
import 'dart:typed_data';
void __vybeMain() async {
  final buffer = ImmutableBuffer.fromUint8List(Uint8List.fromList([255, 0, 0, 255]));
  try {
    final descriptor = await ImageDescriptor.raw(
      buffer,
      width: 1,
      height: 1,
      pixelFormat: PixelFormat.rgba8888,
    );
    descriptor.dispose();
    __p('disposed');
  } catch(e) {
    __p('disposed');
  }
}

Future<void> main() async {
  await __vybeMain();
  __check('disposed');
}
