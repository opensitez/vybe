// vybe-test: dart/interfaces_core/two_classes_implement_same_interface
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

abstract class Speak {
  String say();
}
class Cat implements Speak {
  String say() {
    return 'meow';
  }
}
class Dog implements Speak {
  String say() {
    return 'woof';
  }
}
void __vybeMain() {
  __p(Cat().say());
}

void main() {
  __vybeMain();
  __check('meow');
}
