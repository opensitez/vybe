// vybe-test: dart/switch_statements/int_switch_fallthrough_three_labels_before_break
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
  var score = 3;
  switch (score) {
    case 1:
    case 2:
    case 3:
      __p('pass');
      break;
    default:
      __p('fail');
  }
}

void main() {
  __vybeMain();
  __check('pass');
}
