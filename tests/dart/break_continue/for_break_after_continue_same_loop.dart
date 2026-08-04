// vybe-test: dart/break_continue/for_break_after_continue_same_loop
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
  for (var i = 1; i <= 5; i++) {
    if (i % 2 == 0) continue;
    if (i == 5) break;
    __p(i);
  }
}

void main() {
  __vybeMain();
  __check('1\n3');
}
