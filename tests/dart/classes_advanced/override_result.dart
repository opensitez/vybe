// vybe-test: dart/classes_advanced/override_result
// origin: languages/dart/tests/dart/test_classes_advanced.rs

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

class Animal { String speak() => 'generic'; }
class Dog extends Animal {
  @override
  String speak() => 'woof';
}
void __vybeMain() { var d = Dog(); __p(d.speak()); }

void main() {
  __vybeMain();
  __check('woof');
}
