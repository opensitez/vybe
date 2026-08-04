// vybe-test: dart/break_continue/nested_continue_inner_only_affects_inner
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
  var total = 0;
  for (var i = 0; i < 2; i++) {
    for (var j = 0; j < 4; j++) {
      if (j == 2) continue;
      total++;
    }
  }
  __p(total);
}

void main() {
  __vybeMain();
  __check('6');
}
