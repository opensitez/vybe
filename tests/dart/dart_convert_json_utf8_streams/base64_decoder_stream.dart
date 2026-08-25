// vybe-test: dart/dart_convert_json_utf8_streams/base64_decoder_stream
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
  final stream = Stream.fromIterable(['AQID']);
  final out = await stream.transform(base64.decoder).toList();
  __p('${out[0][0]}:${out[0][1]}');
}

Future<void> main() async {
  await __vybeMain();
  __check('1:2');
}
