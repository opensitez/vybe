// vybe-test: dart/dart_convert_utf8_encode_decode/utf8_encode_emoji
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
  // '🚀' is 4 bytes: 0xF0 0x9F 0x9A 0x80 (240, 159, 154, 128)
  final bytes = utf8.encode('🚀');
  __p('${bytes[0]}:${bytes[1]}:${bytes[2]}:${bytes[3]}');
}

void main() {
  __vybeMain();
  __check('240:159:154:128');
}
