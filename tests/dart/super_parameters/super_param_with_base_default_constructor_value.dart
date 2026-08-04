// vybe-test: dart/super_parameters/super_param_with_base_default_constructor_value
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
  int count = 0;
  Base(this.count);
}
class Sub extends Base {
  Sub(super.count);
}
void __vybeMain() {
  __p(Sub(99).count);
}

void main() {
  __vybeMain();
  __check('99');
}
