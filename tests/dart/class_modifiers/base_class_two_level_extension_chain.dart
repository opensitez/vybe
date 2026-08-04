// vybe-test: dart/class_modifiers/base_class_two_level_extension_chain
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

base class A {
  int val() {
    return 1;
  }
}
class B extends A {}
class C extends B {
  @override
  int val() {
    return super.val() + 2;
  }
}
void __vybeMain() {
  __p(C().val());
}

void main() {
  __vybeMain();
  __check('3');
}
