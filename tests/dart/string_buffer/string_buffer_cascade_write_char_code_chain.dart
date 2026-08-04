// vybe-test: dart/string_buffer/string_buffer_cascade_write_char_code_chain
// origin: languages/dart/tests/dart/test_string_buffer.rs

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
  var buf = StringBuffer()
    ..writeCharCode(68)
    ..writeCharCode(69)
    ..write('F');
  __p(buf.toString());
}

void main() {
  __vybeMain();
  __check('DEF');
}
