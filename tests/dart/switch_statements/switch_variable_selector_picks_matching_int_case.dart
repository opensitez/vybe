// vybe-test: dart/switch_statements/switch_variable_selector_picks_matching_int_case
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
  var choice = 4;
  var picked = choice * 1;
  switch (picked) {
    case 4:
      __p('four');
      break;
    case 8:
      __p('eight');
      break;
    default:
      __p('other');
  }
}

void main() {
  __vybeMain();
  __check('four');
}
