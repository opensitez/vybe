// vybe-test: dart/mixin_linearization/mixin_field_not_shadowed_by_class_field
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

mixin M {
  int x = 100;
}
class C with M {
  int y = 1;
}
void __vybeMain() {
  __p(C().x + C().y);
}

void main() {
  __vybeMain();
  __check('101');
}
