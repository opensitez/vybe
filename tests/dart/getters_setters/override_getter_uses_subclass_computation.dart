// vybe-test: dart/getters_setters/override_getter_uses_subclass_computation
// origin: languages/dart/tests/dart/test_getters_setters.rs

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
  String get sound {
    return '...';
  }
}
class Dog extends Animal {
  String get sound {
    return 'woof';
  }
}
void __vybeMain() {
  __p(Dog().sound);
}

void main() {
  __vybeMain();
  __check('woof');
}
