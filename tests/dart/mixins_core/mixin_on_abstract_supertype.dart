// vybe-test: dart/mixins_core/mixin_on_abstract_supertype
// origin: languages/dart/tests/dart/test_mixins_core.rs

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

abstract class Shape {
  int sides = 0;
}
mixin Polygon on Shape {
  int count() {
    return sides;
  }
}
class Tri extends Shape with Polygon {
  Tri() {
    sides = 3;
  }
}
void __vybeMain() {
  __p(Tri().count());
}

void main() {
  __vybeMain();
  __check('3');
}
