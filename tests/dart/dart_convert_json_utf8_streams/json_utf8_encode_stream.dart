// vybe-test: dart/dart_convert_json_utf8_streams/json_utf8_encode_stream
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
  // Damaged test repaired: the inferred Stream<Map<String, int>> cannot take
  // json.encoder (a StreamTransformer<Object?, String>) under dart 3.10.4 —
  // StreamTransformer is invariant — so the stream must be typed Object?.
  // Also measured: json.encoder as a transformer encodes only ONE top-level
  // value and splits it into several string chunks (6 for this map), so the
  // assertion reassembles the bytes instead of counting events.
  final Stream<Object?> stream = Stream.fromIterable([{'a': 1}]);
  final out = await stream.transform(json.encoder).transform(utf8.encoder).toList();
  __p(utf8.decode(out.expand((e) => e).toList()));
}

Future<void> main() async {
  await __vybeMain();
  __check('{"a":1}');
}
