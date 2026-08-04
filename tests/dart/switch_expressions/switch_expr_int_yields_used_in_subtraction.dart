// vybe-test: dart/switch_expressions/switch_expr_int_yields_used_in_subtraction
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
  var base = 50;
  var delta = switch (3) {
    1 => 5,
    2 => 10,
    3 => 15,
    _ => 0,
  };
  __p(base - delta);
}

void main() {
  __vybeMain();
  __check('35');
}
