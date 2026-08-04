// vybe-test: dart/dart_convert_json_encode_decode/json_encode_with_to_json
// origin: languages/dart/tests/dart/test_dart_convert_json_encode_decode.rs

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
class User {
  final String name;
  User(this.name);
  Map<String, dynamic> toJson() => {'name': name};
}
void __vybeMain() {
  final user = User('Bob');
  __p(jsonEncode(user));
}

void main() {
  __vybeMain();
  __check('{"name":"Bob"}');
}
