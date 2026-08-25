// vybe-test: dart/dart_convert_utf8_encode_decode/utf8_decode_boms_ignored
// origin: languages/dart/tests/dart/test_dart_convert_utf8_encode_decode.rs

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
  // Damaged test repaired: dart 3.10.4 STRIPS a leading UTF-8 BOM in
  // `utf8.decode` (measured — the first code unit is 65, not 0xFEFF); the
  // original expectation asserted the opposite and failed under real dart.
  final str = utf8.decode([0xEF, 0xBB, 0xBF, 65]);
  __p(str.codeUnitAt(0) == 65 && str.length == 1);
}

void main() {
  __vybeMain();
  __check('true');
}
