// vybe-test: dart/loops/for_loop_modulo_pattern_prints_fizz_flags
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
  for (var i = 1; i <= 5; i++) {
    __p(i % 3 == 0);
  }
}

void main() {
  __vybeMain();
  __check('false\nfalse\ntrue\nfalse\nfalse');
}
