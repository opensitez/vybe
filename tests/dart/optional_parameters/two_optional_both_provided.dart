// vybe-test: dart/optional_parameters/two_optional_both_provided
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

void pair([int a = 1, int b = 2]) {
  __p('$a,$b');
}
void __vybeMain() {
  pair(10, 20);
}

void main() {
  __vybeMain();
  __check('10,20');
}
