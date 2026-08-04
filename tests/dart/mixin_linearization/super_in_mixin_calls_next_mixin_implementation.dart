// vybe-test: dart/mixin_linearization/super_in_mixin_calls_next_mixin_implementation
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

mixin A {
  String run() {
    return 'A';
  }
}
mixin B on Object {
  String run() {
    return super.run() + 'B';
  }
}
class C with A, B {}
void __vybeMain() {
  __p(C().run());
}

void main() {
  __vybeMain();
  __check('AB');
}
