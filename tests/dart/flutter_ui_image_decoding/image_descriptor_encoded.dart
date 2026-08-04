// vybe-test: dart/flutter_ui_image_decoding/image_descriptor_encoded
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
  final buffer = ImmutableBuffer.fromUint8List(Uint8List(10));
  try {
    final descriptor = await ImageDescriptor.encoded(buffer);
    __p(descriptor.width);
  } catch(e) {
    __p('encoded_failed');
  }
}

Future<void> main() async {
  await __vybeMain();
  __check('encoded_failed');
}
