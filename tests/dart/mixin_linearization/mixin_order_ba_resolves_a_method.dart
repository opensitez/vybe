// vybe-test: dart/mixin_linearization/mixin_order_ba_resolves_a_method
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
  String tag() {
    return 'a';
  }
}
mixin B {
  String tag() {
    return 'b';
  }
}
class Y with B, A {}
void __vybeMain() {
  __p(Y().tag());
}

void main() {
  __vybeMain();
  __check('a');
}
