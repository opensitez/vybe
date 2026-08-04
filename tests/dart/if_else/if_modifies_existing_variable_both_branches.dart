// vybe-test: dart/if_else/if_modifies_existing_variable_both_branches
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
  var sign = 'zero';
  var n = -3;
  if (n > 0) {
    sign = 'positive';
  } else if (n < 0) {
    sign = 'negative';
  } else {
    sign = 'zero';
  }
  __p(sign);
}

void main() {
  __vybeMain();
  __check('negative');
}
