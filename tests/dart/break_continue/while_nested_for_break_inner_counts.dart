// vybe-test: dart/break_continue/while_nested_for_break_inner_counts
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
  var w = 0;
  while (w < 2) {
    for (var k = 0; k < 4; k++) {
      if (k == 2) break;
      w++;
    }
    w++;
  }
  __p(w);
}

void main() {
  __vybeMain();
  __check('3');
}
