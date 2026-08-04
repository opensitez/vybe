// vybe-test: dart/dart_convert_base64_encode_decode/base64_encode_large_list
// origin: languages/dart/tests/dart/test_dart_convert_base64_encode_decode.rs

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
  final bytes = List<int>.filled(100, 65); // 'A'
  final encoded = base64Encode(bytes);
  __p(encoded.length); // 100 bytes -> ceil(100/3)*4 = 34 * 4 = 136
}

void main() {
  __vybeMain();
  __check('136');
}
