// vybe-test: dart/optional_parameters/three_trailing_optionals_none_given
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

void seq([int a = 0, int b = 0, int c = 0]) {
  __p('$a$b$c');
}
void __vybeMain() {
  seq();
}

void main() {
  __vybeMain();
  __check('000');
}
