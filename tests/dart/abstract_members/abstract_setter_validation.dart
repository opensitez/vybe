// vybe-test: dart/abstract_members/abstract_setter_validation
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

abstract class Validated {
  set age(int v);
  int get age;
}
class Person extends Validated {
  int _age = 0;
  set age(int v) {
    if (v >= 0) {
      _age = v;
    }
  }
  int get age => _age;
}
void __vybeMain() {
  var p = Person();
  p.age = 30;
  __p(p.age);
}

void main() {
  __vybeMain();
  __check('30');
}
