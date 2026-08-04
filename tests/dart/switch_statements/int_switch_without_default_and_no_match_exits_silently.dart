// vybe-test: dart/switch_statements/int_switch_without_default_and_no_match_exits_silently
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
  var x = 7;
  switch (x) {
    case 1:
      __p('one');
      break;
    case 2:
      __p('two');
      break;
  }
  __p('done');
}

void main() {
  __vybeMain();
  __check('done');
}
