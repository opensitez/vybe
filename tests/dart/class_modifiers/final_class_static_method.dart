// vybe-test: dart/class_modifiers/final_class_static_method
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

final class MathUtil {
  static int doubleIt(int n) {
    return n * 2;
  }
}
void __vybeMain() {
  __p(MathUtil.doubleIt(11));
}

void main() {
  __vybeMain();
  __check('22');
}
