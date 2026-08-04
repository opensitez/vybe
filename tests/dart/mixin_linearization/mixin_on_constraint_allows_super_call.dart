// vybe-test: dart/mixin_linearization/mixin_on_constraint_allows_super_call
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

class Engine {
  int power = 10;
}
mixin Turbo on Engine {
  int boosted() {
    return super.power + 5;
  }
}
class Car extends Engine with Turbo {}
void __vybeMain() {
  __p(Car().boosted());
}

void main() {
  __vybeMain();
  __check('15');
}
