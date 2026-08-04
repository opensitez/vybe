// vybe-test: dart/switch_statements/int_switch_break_prevents_fallthrough_to_next_case
// origin: languages/dart/tests/dart/test_switch_statements.rs

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
  var n = 1;
  switch (n) {
    case 1:
      __p('first');
      break;
    case 2:
      __p('second');
      break;
    default:
      __p('rest');
  }
}

void main() {
  __vybeMain();
  __check('first');
}
