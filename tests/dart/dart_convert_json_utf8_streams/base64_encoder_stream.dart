// vybe-test: dart/dart_convert_json_utf8_streams/base64_encoder_stream
// origin: languages/dart/tests/dart/test_dart_convert_json_utf8_streams.rs

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

import 'dart:convert';
void __vybeMain() async {
  final stream = Stream.fromIterable([ [1, 2, 3] ]);
  final out = await stream.transform(base64.encoder).join();
  __p(out);
}

Future<void> main() async {
  await __vybeMain();
  __check('AQID');
}
