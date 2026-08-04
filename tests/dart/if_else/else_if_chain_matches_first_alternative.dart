// vybe-test: dart/if_else/else_if_chain_matches_first_alternative
// origin: languages/dart/tests/dart/test_if_else.rs

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
  var n = 95;
  if (n >= 90) {
    __p('A');
  } else if (n >= 80) {
    __p('B');
  } else if (n >= 70) {
    __p('C');
  } else {
    __p('F');
  }
}

void main() {
  __vybeMain();
  __check('A');
}
