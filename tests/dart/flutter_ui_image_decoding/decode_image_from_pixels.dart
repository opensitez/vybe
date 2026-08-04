// vybe-test: dart/flutter_ui_image_decoding/decode_image_from_pixels
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
void __vybeMain() {
  final pixels = Uint8List.fromList([0, 0, 0, 255]);
  decodeImageFromPixels(
    pixels,
    1,
    1,
    PixelFormat.rgba8888,
    (image) {
      __p('decoded_pixels');
    },
  );
  __p('called_pixels');
}

void main() {
  __vybeMain();
  __check('called_pixels');
}
