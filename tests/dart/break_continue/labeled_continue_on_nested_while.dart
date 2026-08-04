// vybe-test: dart/break_continue/labeled_continue_on_nested_while
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
  var count = 0;
  outer:
  for (var i = 0; i < 2; i++) {
    var j = 0;
    while (j < 3) {
      j++;
      if (j == 2) continue outer;
      count++;
    }
  }
  __p(count);
}

void main() {
  __vybeMain();
  __check('2');
}
