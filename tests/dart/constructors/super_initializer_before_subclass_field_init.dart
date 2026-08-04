// vybe-test: dart/constructors/super_initializer_before_subclass_field_init
// origin: languages/dart/tests/dart/test_constructors.rs

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
  int a;
  Base(this.a);
}
class Sub extends Base {
  int b;
  Sub(int x, int y) : super(x), b = y;
}
void __vybeMain() {
  var s = Sub(2, 3);
  __p(s.a + s.b);
}

void main() {
  __vybeMain();
  __check('5');
}
