// vybe-test: dart/patterns_core/switch_expr_int_nested_switch
// origin: languages/dart/tests/dart/test_patterns_core.rs

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
  var outer = 1;
  var inner = 2;
  __p(switch (outer) {
    1 => switch (inner) {
      2 => 'nested',
      _ => 'inner-other' },
    _ => 'outer-other' });
}

void main() {
  __vybeMain();
  __check('nested');
}
