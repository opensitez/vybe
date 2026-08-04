// vybe-test: dart/mixin_linearization/mixin_method_calls_superclass_method_via_on
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

class Base {
  int compute() {
    return 5;
  }
}
mixin Double on Base {
  int compute() {
    return super.compute() * 2;
  }
}
class Twice extends Base with Double {}
void __vybeMain() {
  __p(Twice().compute());
}

void main() {
  __vybeMain();
  __check('10');
}
