// vybe-test: dart/functions_core/recursive_power_two_to_fifth
// origin: languages/dart/tests/dart/test_functions_core.rs

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

int power(int base, int exp) {
  if (exp == 0) {
    return 1;
  }
  return base * power(base, exp - 1);
}
void __vybeMain() {
  __p(power(2, 5));
}

void main() {
  __vybeMain();
  __check('32');
}
