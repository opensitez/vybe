// vybe-test: dart/constructors/factory_returns_subclass_instance
// origin: languages/dart/tests/dart/test_constructors.rs

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
  String kind;
  Animal(this.kind);
  factory Animal.dog() {
    return Dog('dog');
  }
}
class Dog extends Animal {
  Dog(String k) : super(k);
}
void __vybeMain() {
  __p(Animal.dog().kind);
}

void main() {
  __vybeMain();
  __check('dog');
}
