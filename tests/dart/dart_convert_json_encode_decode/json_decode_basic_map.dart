// vybe-test: dart/dart_convert_json_encode_decode/json_decode_basic_map
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
void __vybeMain() {
  final String jsonStr = '{"name":"Dart","year":2011}';
  final map = jsonDecode(jsonStr);
  __p('${map['name']}:${map['year']}');
}

void main() {
  __vybeMain();
  __check('Dart:2011');
}
