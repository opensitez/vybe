// vybe-test: dart/mixin_linearization/super_in_on_mixin_delegates_to_extended_class
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
  int value = 3;
}
mixin Times on Base {
  int value = 99;
  int read() {
    return super.value;
  }
}
class Box extends Base with Times {}
void __vybeMain() {
  __p(Box().read());
}

void main() {
  __vybeMain();
  __check('3');
}
