// vybe-test: dart/operator_overloading/operator_plus_adds_two_int_fields
// origin: languages/dart/tests/dart/test_operator_overloading.rs

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

class Vec2 {
  int x;
  int y;
  Vec2(this.x, this.y);
  Vec2 operator +(Vec2 other) {
    return Vec2(x + other.x, y + other.y);
  }
}
void __vybeMain() {
  var a = Vec2(1, 2);
  var b = Vec2(3, 4);
  var c = a + b;
  __p(c.x + c.y);
}

void main() {
  __vybeMain();
  __check('10');
}
