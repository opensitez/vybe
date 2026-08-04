// vybe-test: dart/break_continue/for_in_break_on_target_string_char
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
  var hit = '';
  for (var ch in 'abcde') {
    if (ch == 'c') {
      hit = ch;
      break;
    }
  }
  __p(hit);
}

void main() {
  __vybeMain();
  __check('c');
}
