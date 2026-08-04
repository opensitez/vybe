// vybe-test: dart/class_modifiers/mixin_class_on_constraint_with_supertype
// origin: languages/dart/tests/dart/test_class_modifiers.rs

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

class Animal {
  String kind = 'animal';
}
mixin class Pet on Animal {
  String care() {
    return 'feed';
  }
}
class Dog extends Animal with Pet {}
void __vybeMain() {
  __p(Dog().care());
}

void main() {
  __vybeMain();
  __check('feed');
}
