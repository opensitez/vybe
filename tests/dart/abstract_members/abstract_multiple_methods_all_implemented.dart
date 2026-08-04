// vybe-test: dart/abstract_members/abstract_multiple_methods_all_implemented
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

abstract class Calc {
  int add(int a, int b);
  int sub(int a, int b);
}
class BasicCalc extends Calc {
  int add(int a, int b) {
    return a + b;
  }
  int sub(int a, int b) {
    return a - b;
  }
}
void __vybeMain() {
  __p(BasicCalc().add(7, 3));
}

void main() {
  __vybeMain();
  __check('10');
}
