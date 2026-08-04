// vybe-test: dart/mixin_linearization/mixin_super_invokes_parent_mixin_method
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

mixin P {
  String step() {
    return 'P';
  }
}
mixin Q on Object {
  String step() {
    return super.step() + 'Q';
  }
}
class R with P, Q {}
void __vybeMain() {
  __p(R().step());
}

void main() {
  __vybeMain();
  __check('PQ');
}
