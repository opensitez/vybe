// vybe-test: dart/mixin_linearization/mixin_getter_overrides_class_getter
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
  int get val {
    return 1;
  }
}
mixin Override {
  int get val {
    return 9;
  }
}
class Sub extends Base with Override {}
void __vybeMain() {
  __p(Sub().val);
}

void main() {
  __vybeMain();
  __check('9');
}
