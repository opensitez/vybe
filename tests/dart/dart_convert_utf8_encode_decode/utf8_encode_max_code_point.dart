// vybe-test: dart/dart_convert_utf8_encode_decode/utf8_encode_max_code_point
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
  // Max valid Unicode code point is U+10FFFF
  // High surrogate: 0xDBFF, Low surrogate: 0xDFFF
  final str = String.fromCharCodes([0xDBFF, 0xDFFF]);
  final bytes = utf8.encode(str);
  // It takes 4 bytes in UTF-8
  __p(bytes.length == 4);
}

void main() {
  __vybeMain();
  __check('true');
}
