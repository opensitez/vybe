// vybe-test: dart/super_parameters/super_param_plus_subclass_field
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
  int x;
  Base(this.x);
}
class Sub extends Base {
  int y;
  Sub(super.x, this.y);
}
void __vybeMain() {
  var s = Sub(4, 6);
  __p(s.x + s.y);
}

void main() {
  __vybeMain();
  __check('10');
}
