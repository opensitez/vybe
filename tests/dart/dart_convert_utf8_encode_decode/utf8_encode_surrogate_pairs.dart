// vybe-test: dart/dart_convert_utf8_encode_decode/utf8_encode_surrogate_pairs
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
  // High and low surrogates for 🚀 (U+1F680): D83D DE80
  final str = String.fromCharCodes([0xD83D, 0xDE80]);
  final bytes = utf8.encode(str);
  __p(bytes.length); // 4
}

void main() {
  __vybeMain();
  __check('4');
}
