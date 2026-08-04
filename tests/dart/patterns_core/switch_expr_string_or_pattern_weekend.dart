// vybe-test: dart/patterns_core/switch_expr_string_or_pattern_weekend
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
  var day = 'Sat';
  __p(switch (day) {
    'Sat' || 'Sun' => 'weekend',
    _ => 'weekday' });
}

void main() {
  __vybeMain();
  __check('weekend');
}
