// vybe-test: dart/dart_convert_json_utf8_streams/chunked_conversion_sink_json
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

void __vybeMain() {
  var results = [];
  var sink = json.encoder.startChunkedConversion(
    ChunkedConversionSink.withCallback((chunks) {
      results.addAll(chunks);
    })
  );
  sink.add({'a': 1});
  sink.close();
  // Damaged test repaired: dart 3.10.4's chunked JsonEncoder emits the value
  // in MULTIPLE chunks (results[0] measured as just "{"), so the assertion
  // must join the chunks rather than read the first one.
  __p(results.join());
}

void main() {
  __vybeMain();
  __check('{"a":1}');
}
