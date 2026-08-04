// vybe-test: dart/if_else/modulo_in_arithmetic_condition
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
  var n = 10;
  if (n % 3 == 1) {
    __p('remainder-one');
  } else {
    __p('other-remainder');
  }
}

void main() {
  __vybeMain();
  __check('remainder-one');
}
