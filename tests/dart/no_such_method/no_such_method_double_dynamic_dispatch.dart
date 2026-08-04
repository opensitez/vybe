// vybe-test: dart/no_such_method/no_such_method_double_dynamic_dispatch
// origin: languages/dart/tests/dart/test_no_such_method.rs

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
  @override
  dynamic noSuchMethod(Invocation inv) {
    return B();
  }
}
class B {
  int val() {
    return 3;
  }
}
void __vybeMain() {
  dynamic a = A();
  __p(a.next().val());
}

void main() {
  __vybeMain();
  __check('3');
}
