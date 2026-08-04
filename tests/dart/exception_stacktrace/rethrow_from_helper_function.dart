// vybe-test: dart/exception_stacktrace/rethrow_from_helper_function
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

void inner() {
  try {
    throw 'helper';
  } catch (e, st) {
    __p(st != null);
    rethrow;
  }
}
void __vybeMain() {
  try {
    inner();
  } catch (e, st) {
    __p(e);
  }
}

void main() {
  __vybeMain();
  __check('true\nhelper');
}
