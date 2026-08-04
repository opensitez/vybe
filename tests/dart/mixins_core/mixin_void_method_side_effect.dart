// vybe-test: dart/mixins_core/mixin_void_method_side_effect
// origin: languages/dart/tests/dart/test_mixins_core.rs

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

mixin Logger {
  int hits = 0;
  void hit() {
    hits = hits + 1;
  }
}
class Target with Logger {}
void __vybeMain() {
  var t = Target();
  t.hit();
  t.hit();
  __p(t.hits);
}

void main() {
  __vybeMain();
  __check('2');
}
