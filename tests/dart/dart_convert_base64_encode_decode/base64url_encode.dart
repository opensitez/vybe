// vybe-test: dart/dart_convert_base64_encode_decode/base64url_encode
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
  // Bytes that generate '+' and '/' in normal base64: [251, 239]
  // In base64url it should be '-' and '_'
  final bytes = [251, 239]; 
  // base64: ++8=
  // base64url: --8=
  __p(base64UrlEncode(bytes));
}

void main() {
  __vybeMain();
  __check('--8=');
}
