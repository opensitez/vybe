// vybe-test: dart/mixin_linearization/extends_then_with_mixin_order
// origin: languages/dart/tests/dart/test_mixin_linearization.rs

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
  String from() {
    return 'base';
  }
}
mixin Mid {
  String from() {
    return 'mid';
  }
}
class Leaf extends Base with Mid {}
void __vybeMain() {
  __p(Leaf().from());
}

void main() {
  __vybeMain();
  __check('mid');
}
