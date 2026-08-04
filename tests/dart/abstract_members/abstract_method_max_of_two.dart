// vybe-test: dart/abstract_members/abstract_method_max_of_two
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

abstract class Max {
  int max(int a, int b);
}
class IntMax extends Max {
  int max(int a, int b) {
    return a > b ? a : b;
  }
}
void __vybeMain() {
  __p(IntMax().max(12, 7));
}

void main() {
  __vybeMain();
  __check('12');
}
