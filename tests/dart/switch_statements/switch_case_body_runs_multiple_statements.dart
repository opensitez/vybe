// vybe-test: dart/switch_statements/switch_case_body_runs_multiple_statements
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
  var n = 2;
  switch (n) {
    case 2:
      __p('start');
      __p('end');
      break;
    default:
      __p('other');
  }
}

void main() {
  __vybeMain();
  __check('start\nend');
}
