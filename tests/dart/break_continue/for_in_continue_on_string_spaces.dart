// vybe-test: dart/break_continue/for_in_continue_on_string_spaces
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
  var letters = 0;
  for (var ch in 'a b c') {
    if (ch == ' ') continue;
    letters++;
  }
  __p(letters);
}

void main() {
  __vybeMain();
  __check('3');
}
