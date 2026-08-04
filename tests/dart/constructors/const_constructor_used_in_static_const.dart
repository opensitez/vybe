// vybe-test: dart/constructors/const_constructor_used_in_static_const
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

class Vec {
  final int x;
  final int y;
  const Vec(this.x, this.y);
  static const zero = Vec(0, 0);
}
void __vybeMain() {
  __p(Vec.zero.x);
}

void main() {
  __vybeMain();
  __check('0');
}
