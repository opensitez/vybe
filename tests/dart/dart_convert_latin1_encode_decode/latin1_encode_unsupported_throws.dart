// vybe-test: dart/dart_convert_latin1_encode_decode/latin1_encode_unsupported_throws
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
  // '🚀' is U+1F680, not in Latin-1
  try {
    latin1.encode('🚀');
  } on FormatException {
    __p('FormatException thrown');
  } catch(e) {
    __p('ArgumentError thrown'); // Depending on implementation
  }
}

void main() {
  __vybeMain();
  __check('FormatException thrown');
}
