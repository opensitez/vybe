// vybe-test: dart/constructors/const_named_constructor
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

class Origin {
  final int x;
  final int y;
  const Origin(this.x, this.y);
  const Origin.zero() : x = 0, y = 0;
}
void __vybeMain() {
  const o = Origin.zero();
  __p(o.y);
}

void main() {
  __vybeMain();
  __check('0');
}
