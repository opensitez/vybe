// vybe-test: dart/interfaces_core/implements_three_interfaces
// origin: languages/dart/tests/dart/test_interfaces_core.rs

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
abstract class B {
  int b();
}
abstract class C {
  int c();
}
class All implements A, B, C {
  int a() {
    return 1;
  }
  int b() {
    return 2;
  }
  int c() {
    return 3;
  }
}
void __vybeMain() {
  var x = All();
  __p(x.a() + x.b() + x.c());
}

void main() {
  __vybeMain();
  __check('6');
}
