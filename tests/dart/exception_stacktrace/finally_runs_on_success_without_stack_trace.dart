// vybe-test: dart/exception_stacktrace/finally_runs_on_success_without_stack_trace
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
    __p('ok');
  } catch (e, st) {
    __p('fail');
  } finally {
    __p('fin');
  }
}

void main() {
  __vybeMain();
  __check('ok\nfin');
}
