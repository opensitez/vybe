// vybe-test: dart/covariant_keyword/covariant_param_narrow_animal_to_cat
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

class Animal {}
class Cat extends Animal {}
class Cage {
  void admit(Animal a) {}
}
class CatCage extends Cage {
  @override
  void admit(covariant Cat c) {}
}
void __vybeMain() {
  var cage = CatCage();
  cage.admit(Cat());
  __p('ok');
}

void main() {
  __vybeMain();
  __check('ok');
}
