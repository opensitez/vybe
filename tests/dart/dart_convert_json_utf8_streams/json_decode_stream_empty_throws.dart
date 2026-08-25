// vybe-test: dart/dart_convert_json_utf8_streams/json_decode_stream_empty_throws
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
  final stream = Stream<String>.empty();
  try {
    await stream.transform(json.decoder).first;
  } catch(e) {
    __p('StateError thrown'); // first on empty stream throws StateError
  }
}

Future<void> main() async {
  await __vybeMain();
  __check('StateError thrown');
}
