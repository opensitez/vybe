// vybe-test: dart/errors_advanced/finally_always_runs
// origin: languages/dart/tests/dart/test_errors_advanced.rs

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
  try {
    __p('try');
  } finally {
    __p('finally');
  }
}

void main() {
  __vybeMain();
  __check('try\nfinally');
}
