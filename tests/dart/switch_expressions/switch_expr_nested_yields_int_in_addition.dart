// vybe-test: dart/switch_expressions/switch_expr_nested_yields_int_in_addition
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
  var a = 1;
  var b = 2;
  var total = switch (a) {
    1 => switch (b) {
      2 => 10,
      _ => 1,
    },
    _ => 0,
  } + switch (b) {
    2 => 20,
    _ => 0,
  };
  __p(total);
}

void main() {
  __vybeMain();
  __check('30');
}
