// vybe-test: dart/dart_convert_json_utf8_streams/chunked_conversion_sink_utf8
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
void __vybeMain() {
  var results = [];
  var sink = utf8.encoder.startChunkedConversion(
    ByteConversionSink.withCallback((bytes) {
      results.addAll(bytes);
    })
  );
  sink.add('A');
  sink.close();
  __p(results[0]);
}

void main() {
  __vybeMain();
  __check('65');
}
