// vybe-test: dart/switch_expressions/switch_expr_yields_int_in_list_element
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
  var code = 2;
  var list = [
    switch (code) {
      1 => 10,
      2 => 20,
      _ => 0,
    },
    99,
  ];
  __p(list[0]);
  __p(list[1]);
}

void main() {
  __vybeMain();
  __check('20\n99');
}
