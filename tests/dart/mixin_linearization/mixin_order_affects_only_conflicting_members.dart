// vybe-test: dart/mixin_linearization/mixin_order_affects_only_conflicting_members
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

mixin Alpha {
  int a() {
    return 1;
  }
  int clash() {
    return 10;
  }
}
mixin Beta {
  int b() {
    return 2;
  }
  int clash() {
    return 20;
  }
}
class Gamma with Alpha, Beta {}
void __vybeMain() {
  var g = Gamma();
  __p(g.a() + g.b() + g.clash());
}

void main() {
  __vybeMain();
  __check('23');
}
