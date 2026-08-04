// vybe-test: dart/exceptions_core/catch_exception_from_loop
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
    for (var i = 0; i < 3; i++) {
      if (i == 2) throw 'loop break';
    }
  } catch (e) {
    __p(e);
  }
}

void main() {
  __vybeMain();
  __check('loop break');
}
