// vybe-test: dart/classes_core/two_level_inheritance_reuses_grandparent_method
// origin: languages/dart/tests/dart/test_classes_core.rs

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
  int base() {
    return 1;
  }
}
class B extends A {}
class C extends B {}
void __vybeMain() {
  __p(C().base());
}

void main() {
  __vybeMain();
  __check('1');
}
