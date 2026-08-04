// vybe-test: dart/exception_stacktrace/stack_trace_on_custom_exception_class
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

class AppError implements Exception {
  final String msg;
  AppError(this.msg);
}
void __vybeMain() {
  try {
    throw AppError('custom');
  } catch (e, st) {
    __p(st != null);
  }
}

void main() {
  __vybeMain();
  __check('true');
}
