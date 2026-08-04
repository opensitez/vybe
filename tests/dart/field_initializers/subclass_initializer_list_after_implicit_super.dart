// vybe-test: dart/field_initializers/subclass_initializer_list_after_implicit_super
// origin: languages/dart/tests/dart/test_field_initializers.rs

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
  int x = 10;
}
class Sub extends Base {
  int y;
  Sub(int v) : y = v;
}
void __vybeMain() {
  __p(Sub(3).x + Sub(3).y);
}

void main() {
  __vybeMain();
  __check('13');
}
