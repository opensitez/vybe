// vybe-test: dart/closures/make_multiplier_closure_captures_factor
// origin: languages/dart/tests/dart/test_closures.rs

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

int Function(int) makeMultiplier(int m) {
  return (x) => x * m;
}
void __vybeMain() {
  var triple = makeMultiplier(3);
  __p(triple(4));
}

void main() {
  __vybeMain();
  __check('12');
}
