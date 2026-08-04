// vybe-test: dart/abstract_members/abstract_three_level_inheritance
// origin: languages/dart/tests/dart/test_abstract_members.rs

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

abstract class A {
  int a();
}
abstract class B extends A {
  int b();
}
class C extends B {
  int a() {
    return 1;
  }
  int b() {
    return 2;
  }
}
void __vybeMain() {
  __p(C().a() + C().b());
}

void main() {
  __vybeMain();
  __check('3');
}
