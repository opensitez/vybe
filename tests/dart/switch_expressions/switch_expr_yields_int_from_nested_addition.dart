// vybe-test: dart/switch_expressions/switch_expr_yields_int_from_nested_addition
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
  var base = 2;
  var total = switch (base) {
    1 => 1 + 2 + 3,
    2 => 4 + 5 + 6,
    _ => 0,
  };
  __p(total);
}

void main() {
  __vybeMain();
  __check('15');
}
