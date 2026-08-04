// vybe-test: dart/mixins_core/mixin_on_hierarchy_three_levels
// origin: languages/dart/tests/dart/test_mixins_core.rs

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

class A {
  int a() {
    return 1;
  }
}
class B extends A {}
mixin C on B {
  int c() {
    return 10;
  }
}
class D extends B with C {}
void __vybeMain() {
  var d = D();
  __p(d.a() + d.c());
}

void main() {
  __vybeMain();
  __check('11');
}
