// vybe-test: dart/typedefs_core/typedef_double_binary_multiplies
// origin: languages/dart/tests/dart/test_typedefs_core.rs

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

typedef Scale = double Function(double, double);
double mul(double a, double b) {
  return a * b;
}
void __vybeMain() {
  Scale fn = mul;
  __p(fn(2.5, 4.0));
}

void main() {
  __vybeMain();
  __check('10.0');
}
