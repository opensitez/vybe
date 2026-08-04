// vybe-test: dart/switch_expressions/switch_expr_yields_string_uppercase_in_arm
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
  var mode = 'run';
  var banner = switch (mode) {
    'run' => 'RUNNING',
    'stop' => 'STOPPED',
    _ => 'IDLE',
  };
  __p(banner);
}

void main() {
  __vybeMain();
  __check('RUNNING');
}
