// vybe-test: dart/super_parameters/super_param_named_optional_uses_base_default
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
  int level;
  Base({this.level = 5});
}
class Sub extends Base {
  Sub({super.level});
}
void __vybeMain() {
  __p(Sub().level);
}

void main() {
  __vybeMain();
  __check('5');
}
