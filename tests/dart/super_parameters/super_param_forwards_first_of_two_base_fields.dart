// vybe-test: dart/super_parameters/super_param_forwards_first_of_two_base_fields
// origin: languages/dart/tests/dart/test_super_parameters.rs

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
  int b;
  Base(this.a, this.b);
}
class Sub extends Base {
  Sub(super.a, super.b);
}
void __vybeMain() {
  var s = Sub(2, 3);
  __p(s.a + s.b);
}

void main() {
  __vybeMain();
  __check('5');
}
