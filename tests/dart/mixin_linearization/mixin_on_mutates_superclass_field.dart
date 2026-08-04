// vybe-test: dart/mixin_linearization/mixin_on_mutates_superclass_field
// origin: languages/dart/tests/dart/test_mixin_linearization.rs

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
  int n = 0;
}
mixin Bump on Counter {
  void inc() {
    n = n + 1;
  }
}
class Tally extends Counter with Bump {}
void __vybeMain() {
  var t = Tally();
  t.inc();
  __p(t.n);
}

void main() {
  __vybeMain();
  __check('1');
}
