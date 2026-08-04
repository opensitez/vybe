// vybe-test: dart/switch_statements/switch_string_variable_selects_weekday
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
  var day = 'Wed';
  switch (day) {
    case 'Mon':
      __p('start');
      break;
    case 'Wed':
      __p('midweek');
      break;
    case 'Fri':
      __p('end');
      break;
    default:
      __p('other');
  }
}

void main() {
  __vybeMain();
  __check('midweek');
}
