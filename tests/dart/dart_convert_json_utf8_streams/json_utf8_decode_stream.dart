// vybe-test: dart/dart_convert_json_utf8_streams/json_utf8_decode_stream
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
  final stream = Stream.fromIterable([utf8.encode('{"a":1}')]);
  final out = await stream.transform(utf8.decoder).transform(json.decoder).toList();
  __p(out[0]['a']);
}

Future<void> main() async {
  await __vybeMain();
  __check('1');
}
