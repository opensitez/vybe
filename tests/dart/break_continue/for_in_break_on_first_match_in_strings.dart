// vybe-test: dart/break_continue/for_in_break_on_first_match_in_strings
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
  var found = false;
  for (var ch in 'abracadabra') {
    if (ch == 'c') {
      found = true;
      break;
    }
  }
  __p(found);
}

void main() {
  __vybeMain();
  __check('true');
}
