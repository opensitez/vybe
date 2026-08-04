// vybe-test: dart/classes_core/super_method_call_extends_base_result
// origin: languages/dart/tests/dart/test_classes_core.rs

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

class A {
  String greet() {
    return 'hi';
  }
}
class B extends A {
  String greet() {
    return super.greet() + '!';
  }
}
void __vybeMain() {
  __p(B().greet());
}

void main() {
  __vybeMain();
  __check('hi!');
}
