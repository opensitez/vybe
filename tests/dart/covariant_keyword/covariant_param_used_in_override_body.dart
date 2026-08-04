// vybe-test: dart/covariant_keyword/covariant_param_used_in_override_body
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

class Animal {
  String name;
  Animal(this.name);
}
class Dog extends Animal {
  Dog(String n) : super(n);
}
class Trainer {
  void train(Animal a) {}
}
class DogTrainer extends Trainer {
  @override
  void train(covariant Dog d) {
    __p(d.name);
  }
}
void __vybeMain() {
  DogTrainer().train(Dog('Rex'));
}

void main() {
  __vybeMain();
  __check('Rex');
}
