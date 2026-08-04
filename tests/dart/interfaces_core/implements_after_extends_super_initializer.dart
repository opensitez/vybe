// vybe-test: dart/interfaces_core/implements_after_extends_super_initializer
// origin: languages/dart/tests/dart/test_interfaces_core.rs

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

class Base {
  int x;
  Base(this.x);
}
abstract class HasY {
  int y();
}
class Pair extends Base implements HasY {
  int _y;
  Pair(int a, int b) : super(a), _y = b;
  int y() {
    return _y;
  }
}
void __vybeMain() {
  var p = Pair(1, 2);
  __p(p.x + p.y());
}

void main() {
  __vybeMain();
  __check('3');
}
