// vybe-test: dart/callable_objects/call_tear_off_same_instance
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

class Id {
  int call(int n) {
    return n;
  }
}
void __vybeMain() {
  var i = Id();
  var f1 = i.call;
  var f2 = i.call;
  __p(f1(7) == f2(7));
}

void main() {
  __vybeMain();
  __check('true');
}
