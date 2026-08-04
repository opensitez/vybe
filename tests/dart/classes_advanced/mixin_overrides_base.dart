// vybe-test: dart/classes_advanced/mixin_overrides_base
// origin: languages/dart/tests/dart/test_classes_advanced.rs

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

class Base { String greet() => 'base'; }
mixin Override { String greet() => 'mixin'; }
class Child extends Base with Override {}
void __vybeMain() { __p(Child().greet()); }

void main() {
  __vybeMain();
  __check('mixin');
}
