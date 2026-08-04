// vybe-test: dart/static_members/static_method_mutates_static_field
// origin: languages/dart/tests/dart/test_static_members.rs

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

class Counter {
  static int n = 0;
  static void inc() {
    n = n + 1;
  }
}
void __vybeMain() {
  Counter.inc();
  Counter.inc();
  __p(Counter.n);
}

void main() {
  __vybeMain();
  __check('2');
}
