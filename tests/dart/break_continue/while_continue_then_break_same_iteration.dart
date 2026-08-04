// vybe-test: dart/break_continue/while_continue_then_break_same_iteration
// origin: languages/dart/tests/dart/test_break_continue.rs

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
  var i = 0;
  var printed = 0;
  while (i < 5) {
    i++;
    if (i == 2) continue;
    if (i == 4) break;
    printed++;
  }
  __p(printed);
}

void main() {
  __vybeMain();
  __check('2');
}
