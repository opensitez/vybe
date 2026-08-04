// vybe-test: dart/field_initializers/subclass_field_init_after_super_default_constructor
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
  int a = 1;
}
class Sub extends Base {
  int b = 2;
}
void __vybeMain() {
  var s = Sub();
  __p(s.a + s.b);
}

void main() {
  __vybeMain();
  __check('3');
}
