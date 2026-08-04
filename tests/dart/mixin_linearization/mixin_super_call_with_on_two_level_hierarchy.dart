// vybe-test: dart/mixin_linearization/mixin_super_call_with_on_two_level_hierarchy
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

class Grand {
  int g() {
    return 1;
  }
}
class Parent extends Grand {}
mixin Child on Parent {
  int total() {
    return super.g() + 5;
  }
}
class Leaf extends Parent with Child {}
void __vybeMain() {
  __p(Leaf().total());
}

void main() {
  __vybeMain();
  __check('6');
}
