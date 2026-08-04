// vybe-test: dart/switch_statements/switch_fallthrough_prints_shared_message_once
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
  var level = 1;
  switch (level) {
    case 1:
    case 2:
      __p('warn');
      break;
    case 3:
      __p('error');
      break;
    default:
      __p('info');
  }
}

void main() {
  __vybeMain();
  __check('warn');
}
