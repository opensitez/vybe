// vybe-test: dart/dart_convert_base64_encode_decode/base64url_encode_no_padding
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
  // In Dart, you can strip padding manually, or sometimes url encode might not need it, 
  // but base64UrlEncode does add padding by default.
  final bytes = [104, 101, 108, 108, 111]; // "hello" -> aGVsbG8=
  __p(base64UrlEncode(bytes).endsWith('='));
}

void main() {
  __vybeMain();
  __check('true');
}
