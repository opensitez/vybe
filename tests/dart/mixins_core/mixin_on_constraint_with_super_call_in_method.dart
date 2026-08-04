// vybe-test: dart/mixins_core/mixin_on_constraint_with_super_call_in_method
// origin: languages/dart/tests/dart/test_mixins_core.rs

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
  int n = 1;
  int baseInc() {
    return n + 1;
  }
}
mixin Wrap on Base {
  int wrapped() {
    return baseInc() + 1;
  }
}
class Sub extends Base with Wrap {}
void __vybeMain() {
  __p(Sub().wrapped());
}

void main() {
  __vybeMain();
  __check('3');
}
