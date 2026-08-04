// vybe-test: dart/switch_expressions/switch_expr_int_yields_max_of_two_branches
// origin: languages/dart/tests/dart/test_switch_expressions.rs

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

void __vybeMain() {
  var pick = switch (2) {
    1 => 3,
    2 => 9,
    _ => 0,
  };
  var other = switch (1) {
    1 => 5,
    _ => 0,
  };
  __p(pick > other ? pick : other);
}

void main() {
  __vybeMain();
  __check('9');
}
