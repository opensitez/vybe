// vybe-test: dart/super_calls/super_in_three_level_chain
// origin: languages/dart/tests/dart/test_super_calls.rs

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
  int depth() {
    return 1;
  }
}
class B extends A {
  int depth() {
    return super.depth() + 1;
  }
}
class C extends B {
  int depth() {
    return super.depth() + 1;
  }
}
void __vybeMain() {
  __p(C().depth());
}

void main() {
  __vybeMain();
  __check('3');
}
