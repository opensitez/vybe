// vybe-test: dart/interfaces_core/subclass_implements_parent_interface
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

abstract class I {
  int get n;
}
class Base implements I {
  int get n {
    return 1;
  }
}
class Sub extends Base {
  int get n {
    return 2;
  }
}
void __vybeMain() {
  __p(Sub().n);
}

void main() {
  __vybeMain();
  __check('2');
}
