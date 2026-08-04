// vybe-test: dart/callable_objects/call_stored_in_variable_and_reused
// origin: languages/dart/tests/dart/test_callable_objects.rs

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

class Mul {
  int call(int a, int b) {
    return a * b;
  }
}
void __vybeMain() {
  var m = Mul();
  __p(m(2, 3) + m(4, 5));
}

void main() {
  __vybeMain();
  __check('26');
}
