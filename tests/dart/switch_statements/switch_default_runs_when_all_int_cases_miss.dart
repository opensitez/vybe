// vybe-test: dart/switch_statements/switch_default_runs_when_all_int_cases_miss
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
  var port = 8080;
  switch (port) {
    case 80:
      __p('http');
      break;
    case 443:
      __p('https');
      break;
    default:
      __p('custom');
  }
}

void main() {
  __vybeMain();
  __check('custom');
}
