// vybe-test: dart/super_parameters/super_param_two_named_both_required
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
  Base({required this.a, required this.b});
}
class Sub extends Base {
  Sub({required super.a, required super.b});
}
void __vybeMain() {
  __p(Sub(a: 3, b: 4).a + Sub(a: 3, b: 4).b);
}

void main() {
  __vybeMain();
  __check('7');
}
