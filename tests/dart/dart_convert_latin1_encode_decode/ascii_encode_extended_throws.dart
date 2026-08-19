// vybe-test: dart/dart_convert_latin1_encode_decode/ascii_encode_extended_throws
// origin: languages/dart/tests/dart/test_dart_convert_latin1_encode_decode.rs

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
  // 'ä' is 228, out of ASCII range (0-127)
  try {
    ascii.encode('ä');
  } on FormatException {
    __p('FormatException thrown');
  }
}

void main() {
  __vybeMain();
  __check('FormatException thrown');
}
