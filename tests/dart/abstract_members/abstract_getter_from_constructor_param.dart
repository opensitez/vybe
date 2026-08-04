// vybe-test: dart/abstract_members/abstract_getter_from_constructor_param
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

abstract class Named {
  String get name;
}
class Person extends Named {
  final String _name;
  Person(this._name);
  String get name => _name;
}
void __vybeMain() {
  __p(Person('Eve').name);
}

void main() {
  __vybeMain();
  __check('Eve');
}
