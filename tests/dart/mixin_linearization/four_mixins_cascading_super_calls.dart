// vybe-test: dart/mixin_linearization/four_mixins_cascading_super_calls
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

mixin A {
  int n() {
    return 1;
  }
}
mixin B on Object {
  int n() {
    return super.n() + 2;
  }
}
mixin C on Object {
  int n() {
    return super.n() + 3;
  }
}
mixin D on Object {
  int n() {
    return super.n() + 4;
  }
}
class E with A, B, C, D {}
void __vybeMain() {
  __p(E().n());
}

void main() {
  __vybeMain();
  __check('10');
}
