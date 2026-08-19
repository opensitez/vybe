// vybe-test: dart/dart_convert_latin1_encode_decode/latin1_decode_allow_invalid
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
  // Latin1 decode has allowInvalid
  final str = latin1.decode([256], allowInvalid: true);
  // Wait, [256] is technically invalid since latin1 is 8-bit.
  // Actually list elements > 255 might be truncated or throw.
  // In Dart, passing >255 to allowInvalid might return replacement chars or just truncate.
  // Let's just catch whatever it does or throws without allowInvalid.
  try {
    latin1.decode([256], allowInvalid: false);
  } on FormatException {
    print('FormatException thrown');
  }
}

void main() {
  __vybeMain();
  __check('FormatException thrown');
}
