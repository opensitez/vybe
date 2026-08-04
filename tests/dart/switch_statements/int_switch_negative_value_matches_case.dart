// vybe-test: dart/switch_statements/int_switch_negative_value_matches_case
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
  var delta = -1;
  switch (delta) {
    case -1:
      __p('minus-one');
      break;
    case 0:
      __p('zero');
      break;
    default:
      __p('positive');
  }
}

void main() {
  __vybeMain();
  __check('minus-one');
}
