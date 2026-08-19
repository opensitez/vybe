// vybe-test: dart/dart_convert_json_utf8_streams/json_decode_stream_multiple_objects_fails
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

void __vybeMain() async {
  final stream = Stream.fromIterable(['{"a":1}{"b":2}']);
  try {
    await stream.transform(json.decoder).toList();
  } on FormatException {
    __p('FormatException thrown');
  }
}

Future<void> main() async {
  await __vybeMain();
  __check('FormatException thrown');
}
