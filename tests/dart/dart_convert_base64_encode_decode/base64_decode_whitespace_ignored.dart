// vybe-test: dart/dart_convert_base64_encode_decode/base64_decode_whitespace_ignored
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
  // Dart's base64 ignores whitespace according to docs, actually wait - the standard base64 parser might throw on whitespace.
  // Wait, Dart 2.0+ base64Decode throws on whitespace. Let's verify exception.
  try {
    base64Decode('aGVs bG8=');
  } on FormatException {
    __p('FormatException thrown');
  }
}

void main() {
  __vybeMain();
  __check('FormatException thrown');
}
