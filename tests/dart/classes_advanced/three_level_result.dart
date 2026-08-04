// vybe-test: dart/classes_advanced/three_level_result
// origin: languages/dart/tests/dart/test_classes_advanced.rs

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

class A { int val() => 1; }
class B extends A { int bonus() => 10; }
class C extends B {}
void __vybeMain() { var c = C(); __p(c.val() + c.bonus()); }

void main() {
  __vybeMain();
  __check('11');
}
