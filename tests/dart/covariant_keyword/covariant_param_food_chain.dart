// vybe-test: dart/covariant_keyword/covariant_param_food_chain
// origin: languages/dart/tests/dart/test_covariant_keyword.rs

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

class Food {}
class Fruit extends Food {}
class Apple extends Fruit {}
class Eater {
  void eat(Food f) {}
}
class FruitEater extends Eater {
  @override
  void eat(covariant Fruit f) {}
}
class AppleEater extends FruitEater {
  @override
  void eat(covariant Apple a) {
    __p('apple');
  }
}
void __vybeMain() {
  AppleEater().eat(Apple());
}

void main() {
  __vybeMain();
  __check('apple');
}
