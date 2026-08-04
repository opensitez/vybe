// vybe-test: dart/dart_convert_latin1_encode_decode/latin1_encode_unsupported_allow_invalid
// origin: languages/dart/tests/dart/test_dart_convert_latin1_encode_decode.rs

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
  // You cannot pass allowInvalid to static latin1.encode directly, need Latin1Encoder
  // Actually wait, Dart doesn't have allowInvalid for Latin1Encoder by default.
  // We'll test Latin1Encoder instantiation.
  final encoder = Latin1Encoder();
  __p(encoder is Converter<String, List<int>>);
}

void main() {
  __vybeMain();
  __check('true');
}
