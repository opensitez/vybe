// vybe-test: dart/dart_convert_utf8_encode_decode/utf8_encode_multibyte
// origin: languages/dart/tests/dart/test_dart_convert_utf8_encode_decode.rs

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
  // 'ä' is 2 bytes: 0xC3 0xA4 (195, 164)
  final bytes = utf8.encode('ä');
  __p('${bytes[0]}:${bytes[1]}');
}

void main() {
  __vybeMain();
  __check('195:164');
}
