// vybe-test: dart/string_buffer/string_buffer_clear_empties_content
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
  var buf = StringBuffer();
  buf.write('remove me');
  buf.clear();
  __p(buf.toString());
  __p(buf.length);
  __p(buf.isEmpty);
}

void main() {
  __vybeMain();
  __check('\n0\ntrue');
}
