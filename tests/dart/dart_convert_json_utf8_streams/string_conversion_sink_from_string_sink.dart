// vybe-test: dart/dart_convert_json_utf8_streams/string_conversion_sink_from_string_sink
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
  var buffer = StringBuffer();
  var sink = StringConversionSink.fromStringSink(buffer);
  sink.add('X');
  sink.add('Y');
  sink.close();
  __p(buffer.toString());
}

void main() {
  __vybeMain();
  __check('XY');
}
