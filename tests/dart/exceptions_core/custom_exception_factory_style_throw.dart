// vybe-test: dart/exceptions_core/custom_exception_factory_style_throw
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

class ServiceError implements Exception {
  final String service;
  ServiceError(this.service);
}
void __vybeMain() {
  try {
    throw ServiceError('auth');
  } catch (e) {
    var s = e as ServiceError;
    __p(s.service);
  }
}

void main() {
  __vybeMain();
  __check('auth');
}
