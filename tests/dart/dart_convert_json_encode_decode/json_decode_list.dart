// vybe-test: dart/dart_convert_json_encode_decode/json_decode_list
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
  final list = jsonDecode('[1,"two",false,null]');
  __p('${list[0]}:${list[1]}:${list[2]}:${list[3]}');
}

void main() {
  __vybeMain();
  __check('1:two:false:null');
}
