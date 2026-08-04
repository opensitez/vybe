// vybe-test: dart/classes_core/override_method_replaces_superclass_behavior
// origin: languages/dart/tests/dart/test_classes_core.rs

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
