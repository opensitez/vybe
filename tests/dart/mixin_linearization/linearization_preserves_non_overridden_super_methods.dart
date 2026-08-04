// vybe-test: dart/mixin_linearization/linearization_preserves_non_overridden_super_methods
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
  int baseOnly() {
    return 7;
  }
}
mixin M {
  int mixOnly() {
    return 3;
  }
}
class Both extends Base with M {}
void __vybeMain() {
  var b = Both();
  __p(b.baseOnly() + b.mixOnly());
}

void main() {
  __vybeMain();
  __check('10');
}
