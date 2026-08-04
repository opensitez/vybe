// vybe-test: dart/patterns_core/switch_expr_when_even_odd_guard
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
  __p(switch (6) {
    var n when n % 2 == 0 => 'even',
    var n when n % 2 == 1 => 'odd',
    _ => 'unknown' });
}

void main() {
  __vybeMain();
  __check('even');
}
