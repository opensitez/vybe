// vybe-test: dart/mixin_linearization/mixin_order_change_alters_super_chain_output
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
  String build() {
    return 'a';
  }
}
mixin B on Object {
  String build() {
    return super.build() + 'b';
  }
}
class AB with A, B {}
class BA with B, A {}
void __vybeMain() {
  __p(AB().build().length + BA().build().length);
}

void main() {
  __vybeMain();
  __check('4');
}
