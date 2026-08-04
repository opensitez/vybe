// vybe-test: dart/class_modifiers/mixin_class_multiple_with_on_class
// origin: languages/dart/tests/dart/test_class_modifiers.rs

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

mixin class A {
  int a() {
    return 1;
  }
}
mixin class B {
  int b() {
    return 2;
  }
}
class Both with A, B {}
void __vybeMain() {
  __p(Both().a() + Both().b());
}

void main() {
  __vybeMain();
  __check('3');
}
