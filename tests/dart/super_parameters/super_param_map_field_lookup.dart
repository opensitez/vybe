// vybe-test: dart/super_parameters/super_param_map_field_lookup
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
  Map<String, int> data;
  Base(this.data);
}
class Sub extends Base {
  Sub(super.data);
}
void __vybeMain() {
  __p(Sub({'a': 1}).data['a']);
}

void main() {
  __vybeMain();
  __check('1');
}
