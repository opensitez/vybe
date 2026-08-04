// vybe-test: dart/const_deep/const_constructor_redirecting
// origin: languages/dart/tests/dart/test_const_deep.rs

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

class Counter {
  final int value;
  const Counter(this.value);
  const Counter.zero() : value = 0;
}
void __vybeMain() {
  const c = Counter.zero();
  __p(c.value);
}

void main() {
  __vybeMain();
  __check('0');
}
