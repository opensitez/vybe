// vybe-test: dart/classes_core/subclass_constructor_sets_own_field
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

class Base {
  int a = 0;
}
class Sub extends Base {
  int b = 0;
  Sub(int x) {
    b = x;
  }
}
void __vybeMain() {
  __p(Sub(4).b);
}

void main() {
  __vybeMain();
  __check('4');
}
