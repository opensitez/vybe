// vybe-test: dart/abstract_members/abstract_method_recursive_fib
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

abstract class Seq {
  int at(int n);
}
class Fib extends Seq {
  int at(int n) {
    if (n <= 1) {
      return n;
    }
    return at(n - 1) + at(n - 2);
  }
}
void __vybeMain() {
  __p(Fib().at(6));
}

void main() {
  __vybeMain();
  __check('8');
}
