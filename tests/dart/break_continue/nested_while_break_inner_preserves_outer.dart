// vybe-test: dart/break_continue/nested_while_break_inner_preserves_outer
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
  var outer = 0;
  var r = 0;
  while (r < 2) {
    var c = 0;
    while (c < 4) {
      if (c == 2) break;
      c++;
    }
    outer++;
    r++;
  }
  __p(outer);
}

void main() {
  __vybeMain();
  __check('2');
}
