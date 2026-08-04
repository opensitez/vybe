// vybe-test: dart/exceptions_core/finally_runs_even_when_catch_handles
// origin: languages/dart/tests/dart/test_exceptions_core.rs

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
    throw 'x';
  } catch (e) {
    __p('handled');
  } finally {
    __p('always');
  }
}

void main() {
  __vybeMain();
  __check('handled\nalways');
}
