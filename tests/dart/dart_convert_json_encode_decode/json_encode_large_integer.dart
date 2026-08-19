// vybe-test: dart/dart_convert_json_encode_decode/json_encode_large_integer
// origin: languages/dart/tests/dart/test_dart_convert_json_encode_decode.rs

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
  // Dart ints are 64-bit, JS JSON limits safely to 53-bit. 
  // Dart's jsonEncode on native supports full 64-bit int serialization
  final map = {'big': 9007199254740992}; // 2^53
  print(jsonEncode(map));
}

void main() {
  __vybeMain();
  __check('{"big":9007199254740992}');
}
