// vybe-test: dart/field_initializers/super_initializer_runs_before_subclass_field_init
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
  int baseVal;
  Base(this.baseVal);
}
class Sub extends Base {
  int subVal;
  Sub(int b, int s) : super(b), subVal = s;
}
void __vybeMain() {
  var s = Sub(10, 20);
  __p(s.baseVal + s.subVal);
}

void main() {
  __vybeMain();
  __check('30');
}
