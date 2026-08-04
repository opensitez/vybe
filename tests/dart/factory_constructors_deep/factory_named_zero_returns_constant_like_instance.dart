// vybe-test: dart/factory_constructors_deep/factory_named_zero_returns_constant_like_instance
// origin: languages/dart/tests/dart/test_factory_constructors_deep.rs

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
  int x;
  int y;
  Vec(this.x, this.y);
  factory Vec.zero() {
    return Vec(0, 0);
  }
}
void __vybeMain() {
  __p(Vec.zero().x + Vec.zero().y);
}

void main() {
  __vybeMain();
  __check('0');
}
