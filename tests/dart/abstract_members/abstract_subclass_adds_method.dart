// vybe-test: dart/abstract_members/abstract_subclass_adds_method
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

abstract class Animal {
  String speak();
}
class Dog extends Animal {
  String speak() {
    return 'woof';
  }
  String fetch() {
    return 'ball';
  }
}
void __vybeMain() {
  __p(Dog().speak() + Dog().fetch());
}

void main() {
  __vybeMain();
  __check('woofball');
}
