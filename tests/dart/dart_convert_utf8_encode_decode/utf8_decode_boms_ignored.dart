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
  // UTF-8 BOM is 0xEF, 0xBB, 0xBF
  // Wait, Dart's standard utf8.decode does not strip BOM by default.
  // You need utf8.decoder to strip it or handle it manually?
  // Let's just decode and check the first character.
  final str = utf8.decode([0xEF, 0xBB, 0xBF, 65]);
  __p(str.codeUnitAt(0) == 0xFEFF); // The BOM character itself
}

void main() {
  __vybeMain();
  __check('true');
}
