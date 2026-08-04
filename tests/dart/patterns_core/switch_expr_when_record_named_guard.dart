// vybe-test: dart/patterns_core/switch_expr_when_record_named_guard
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
  var u = (role: 'admin', level: 3);
  __p(switch (u) {
    (role: 'admin', level: var lv) when lv >= 2 => 'elevated',
    _ => 'standard' });
}

void main() {
  __vybeMain();
  __check('elevated');
}
