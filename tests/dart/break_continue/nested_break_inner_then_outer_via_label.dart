// vybe-test: dart/break_continue/nested_break_inner_then_outer_via_label
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
  var printed = 0;
  outer:
  for (var i = 0; i < 2; i++) {
    for (var j = 0; j < 2; j++) {
      if (j == 0) continue;
      printed++;
      break outer;
    }
  }
  __p(printed);
}

void main() {
  __vybeMain();
  __check('1');
}
