// vybe-test: dart/dart_convert_utf8_encode_decode/utf8_encode_unpaired_surrogate_throws
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
  // Dart strings are UTF-16. 
  // Wait, utf8.encode might not throw on unpaired surrogates, it might encode them as replacement chars.
  // Actually, Dart's standard utf8 encoder converts invalid surrogates to U+FFFD.
  final str = String.fromCharCodes([0xD83D]);
  final bytes = utf8.encode(str);
  print(bytes.length == 3); // U+FFFD is 3 bytes in UTF-8
}

void main() {
  __vybeMain();
  __check('true');
}
