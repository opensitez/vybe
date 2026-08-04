// vybe-test: dart/cascades/cascade_string_buffer_write_then_writeln_adds_line_break
// origin: languages/dart/tests/dart/test_cascades.rs

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
  buf..write('line')..writeln('end');
  __p(buf.toString());
}

void main() {
  __vybeMain();
  __check('lineend\n');
}
