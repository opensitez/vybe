// vybe-test: dart/switch_statements/string_switch_break_stops_at_first_matching_case
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
  var mode = 'read';
  switch (mode) {
    case 'read':
      __p('reading');
      break;
    case 'write':
      __p('writing');
      break;
    default:
      __p('idle');
  }
}

void main() {
  __vybeMain();
  __check('reading');
}
