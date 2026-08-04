// vybe-test: dart/callable_objects/call_instance_equality_not_same_as_tear_off
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

class Fn {
  int call(int n) {
    return n;
  }
}
void __vybeMain() {
  var f = Fn();
  __p(f == f.call);
}

void main() {
  __vybeMain();
  __check('false');
}
