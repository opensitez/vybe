// vybe-test: dart/switch_expressions/switch_expr_yields_string_from_variable_selector
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
  var key = 'east';
  var dir = switch (key) {
    'north' => 'N',
    'south' => 'S',
    'east' => 'E',
    'west' => 'W',
    _ => '?',
  };
  __p(dir);
}

void main() {
  __vybeMain();
  __check('E');
}
