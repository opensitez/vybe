// vybe-test: dart/covariant_keyword/covariant_param_three_overrides_in_chain
// origin: languages/dart/tests/dart/test_covariant_keyword.rs

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

class A {}
class B extends A {}
class C extends B {}
class Handler {
  void handle(A a) {}
}
class BHandler extends Handler {
  @override
  void handle(covariant B b) {}
}
class CHandler extends BHandler {
  @override
  void handle(covariant C c) {
    __p('c');
  }
}
void __vybeMain() {
  CHandler().handle(C());
}

void main() {
  __vybeMain();
  __check('c');
}
