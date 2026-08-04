// vybe-test: dart/mixin_linearization/mixin_linearization_instance_per_class
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

mixin Tag {
  String name = 'tagged';
}
class A with Tag {}
class B with Tag {}
void __vybeMain() {
  var a = A();
  var b = B();
  b.name = 'changed';
  __p(a.name);
}

void main() {
  __vybeMain();
  __check('tagged');
}
