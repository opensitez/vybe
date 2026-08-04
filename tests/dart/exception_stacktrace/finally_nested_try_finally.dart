// vybe-test: dart/exception_stacktrace/finally_nested_try_finally
// origin: languages/dart/tests/dart/test_exception_stacktrace.rs

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
    try {
      throw 'deep';
    } catch (e, st) {
      __p(st != null);
    } finally {
      __p('inner');
    }
  } finally {
    __p('outer');
  }
}

void main() {
  __vybeMain();
  __check('true\ninner\nouter');
}
