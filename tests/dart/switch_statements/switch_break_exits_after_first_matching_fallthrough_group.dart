// vybe-test: dart/switch_statements/switch_break_exits_after_first_matching_fallthrough_group
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
  var code = 3;
  switch (code) {
    case 1:
    case 2:
      __p('small');
      break;
    case 3:
    case 4:
      __p('medium');
      break;
    default:
      __p('large');
  }
}

void main() {
  __vybeMain();
  __check('medium');
}
