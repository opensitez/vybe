// vybe-test: dart/class_modifiers/final_class_method_returns_computed_value
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

final class Rect {
  int w = 4;
  int h = 5;
  int area() {
    return w * h;
  }
}
void __vybeMain() {
  __p(Rect().area());
}

void main() {
  __vybeMain();
  __check('20');
}
