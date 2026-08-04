// vybe-test: dart/super_parameters/super_param_mid_adds_field_leaf_forwards
// origin: languages/dart/tests/dart/test_super_parameters.rs

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
  int base;
  A(this.base);
}
class B extends A {
  int mid;
  B(super.base, this.mid);
}
class C extends B {
  C(super.base, super.mid);
}
void __vybeMain() {
  __p(C(2, 5).mid);
}

void main() {
  __vybeMain();
  __check('5');
}
