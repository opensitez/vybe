// vybe-test: dart/exceptions_core/custom_exception_extends_implements
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

class DataException implements Exception {
  final int code;
  DataException(this.code);
}
void __vybeMain() {
  try {
    throw DataException(42);
  } catch (e) {
    var d = e as DataException;
    __p(d.code);
  }
}

void main() {
  __vybeMain();
  __check('42');
}
