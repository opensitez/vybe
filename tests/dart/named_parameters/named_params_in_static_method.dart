// vybe-test: dart/named_parameters/named_params_in_static_method
// origin: languages/dart/tests/dart/test_named_parameters.rs

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

class Math {
  static int mul({int a = 1, int b = 1}) => a * b;
}
void __vybeMain() {
  __p(Math.mul(a: 6, b: 7));
}

void main() {
  __vybeMain();
  __check('42');
}
