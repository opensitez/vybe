// vybe-test: dart/constructors/const_constructor_multiple_fields
// origin: languages/dart/tests/dart/test_constructors.rs

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

class Pair {
  final int a;
  final int b;
  const Pair(this.a, this.b);
}
void __vybeMain() {
  const p = Pair(2, 3);
  __p(p.a + p.b);
}

void main() {
  __vybeMain();
  __check('5');
}
