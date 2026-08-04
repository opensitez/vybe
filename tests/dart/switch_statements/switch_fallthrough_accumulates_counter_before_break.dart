// vybe-test: dart/switch_statements/switch_fallthrough_accumulates_counter_before_break
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
  var step = 2;
  var total = 0;
  switch (step) {
    case 1:
      total = total + 1;
    case 2:
      total = total + 2;
    case 3:
      total = total + 3;
      break;
    default:
      total = total + 0;
  }
  __p(total);
}

void main() {
  __vybeMain();
  __check('2');
}
