// vybe-test: dart/dart_convert_json_utf8_streams/string_conversion_sink_as_utf8_sink
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
  var result = '';
  var outSink = StringConversionSink.withCallback((s) => result += s);
  var inSink = outSink.asUtf8Sink(false);
  inSink.add([67, 68]);
  inSink.close();
  __p(result);
}

void main() {
  __vybeMain();
  __check('CD');
}
