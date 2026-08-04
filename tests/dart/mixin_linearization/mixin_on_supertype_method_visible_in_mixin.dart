// vybe-test: dart/mixin_linearization/mixin_on_supertype_method_visible_in_mixin
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

class Parent {
  String greet() {
    return 'hi';
  }
}
mixin Child on Parent {
  String shout() {
    return greet().toUpperCase();
  }
}
class Kid extends Parent with Child {}
void __vybeMain() {
  __p(Kid().shout());
}

void main() {
  __vybeMain();
  __check('HI');
}
