// vybe-test: dart/mixin_linearization/mixin_on_constraint_method_composes_with_class_method
// origin: languages/dart/tests/dart/test_mixin_linearization.rs

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

class Worker {
  int basePay() {
    return 100;
  }
}
mixin Bonus on Worker {
  int totalPay() {
    return basePay() + 50;
  }
}
class Employee extends Worker with Bonus {}
void __vybeMain() {
  __p(Employee().totalPay());
}

void main() {
  __vybeMain();
  __check('150');
}
