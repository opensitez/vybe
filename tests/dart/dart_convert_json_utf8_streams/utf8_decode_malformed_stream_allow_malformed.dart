// vybe-test: dart/dart_convert_json_utf8_streams/utf8_decode_malformed_stream_allow_malformed
// origin: languages/dart/tests/dart/test_dart_convert_json_utf8_streams.rs

import 'dart:convert';

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

Future<void> __vybeMain() async {
  final decoder = Utf8Decoder(allowMalformed: true);
  final stream = Stream.fromIterable([[0xFF]]); // Invalid UTF-8
  final out = await stream.transform(decoder).join();
  __p(out.length > 0); // Contains replacement character
}

Future<void> main() async {
  await __vybeMain();
  __check('true');
}
