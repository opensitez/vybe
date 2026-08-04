// vybe-test: dart/flutter_ui_image_decoding/codec_instantiate_image_codec
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
  final list = Uint8List.fromList([137, 80, 78, 71, 13, 10, 26, 10]); // PNG header
  try {
    // Usually throws if not valid PNG, but we test the API invocation
    final codec = await instantiateImageCodec(list);
    __p(codec.frameCount);
  } catch(e) {
    __p('invalid_image_data');
  }
}

Future<void> main() async {
  await __vybeMain();
  __check('invalid_image_data');
}
