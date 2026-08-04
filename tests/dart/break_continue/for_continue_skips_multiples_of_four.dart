// vybe-test: dart/break_continue/for_continue_skips_multiples_of_four
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
  for (var i = 1; i <= 8; i++) {
    if (i % 4 == 0) continue;
    __p(i);
  }
}

void main() {
  __vybeMain();
  __check('1\n2\n3\n5\n6\n7\n8');
}
