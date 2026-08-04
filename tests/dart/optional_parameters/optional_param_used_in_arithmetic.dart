// vybe-test: dart/optional_parameters/optional_param_used_in_arithmetic
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

int bump([int step = 1]) => 10 + step;
void __vybeMain() {
  __p(bump());
  __p(bump(5));
}

void main() {
  __vybeMain();
  __check('11\n15');
}
