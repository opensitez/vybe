// vybe-test: dart/class_modifiers/base_class_extended_in_same_library
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

base class Animal {
  String speak() {
    return '...';
  }
}
class Dog extends Animal {
  @override
  String speak() {
    return 'woof';
  }
}
void __vybeMain() {
  __p(Dog().speak());
}

void main() {
  __vybeMain();
  __check('woof');
}
