// vybe-test: dart/covariant_keyword/covariant_param_employee_manager
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

class Person {
  String name;
  Person(this.name);
}
class Employee extends Person {
  Employee(String n) : super(n);
}
class Dept {
  void hire(Person p) {}
}
class HR extends Dept {
  @override
  void hire(covariant Employee e) {
    __p(e.name);
  }
}
void __vybeMain() {
  HR().hire(Employee('Sam'));
}

void main() {
  __vybeMain();
  __check('Sam');
}
