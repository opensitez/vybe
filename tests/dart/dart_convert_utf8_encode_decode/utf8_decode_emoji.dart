// vybe-test: dart/dart_convert_utf8_encode_decode/utf8_decode_emoji
// origin: languages/dart/tests/dart/test_dart_convert_utf8_encode_decode.rs

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
  final str = utf8.decode([240, 159, 154, 128]);
  __p(str);
}

void main() {
  __vybeMain();
  __check('🚀');
}
