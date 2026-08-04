// vybe-test: dart/optional_parameters/optional_zero_default_not_confused_with_missing
// origin: languages/dart/tests/dart/test_optional_parameters.rs

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

void show([int n = 0]) {
  __p(n == 0 ? 'zero' : '$n');
}
void __vybeMain() {
  show(0);
}

void main() {
  __vybeMain();
  __check('zero');
}
