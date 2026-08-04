// vybe-test: dart/abstract_members/abstract_multiple_getters
// origin: languages/dart/tests/dart/test_abstract_members.rs

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

abstract class Point2D {
  int get x;
  int get y;
}
class Vec extends Point2D {
  int _x;
  int _y;
  Vec(this._x, this._y);
  int get x => _x;
  int get y => _y;
}
void __vybeMain() {
  __p(Vec(3, 7).x + Vec(3, 7).y);
}

void main() {
  __vybeMain();
  __check('10');
}
