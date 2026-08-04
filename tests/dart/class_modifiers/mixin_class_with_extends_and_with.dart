// vybe-test: dart/class_modifiers/mixin_class_with_extends_and_with
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

class Base {
  int baseVal() {
    return 1;
  }
}
mixin class Extra {
  int extraVal() {
    return 2;
  }
}
class Combined extends Base with Extra {}
void __vybeMain() {
  var c = Combined();
  __p(c.baseVal() + c.extraVal());
}

void main() {
  __vybeMain();
  __check('3');
}
