// vybe-test: dart/loops/for_loop_builds_comma_separated_string
// origin: languages/dart/tests/dart/test_loops.rs

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
  var parts = <String>[];
  for (var i = 0; i < 3; i++) {
    parts.add('x$i');
  }
  __p(parts.join(','));
}

void main() {
  __vybeMain();
  __check('x0,x1,x2');
}
