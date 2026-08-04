// vybe-test: dart/mixins_core/mixin_method_can_mutate_mixin_field
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

mixin Counter {
  int n = 0;
  void bump() {
    n = n + 1;
  }
}
class Box with Counter {}
void __vybeMain() {
  var b = Box();
  b.bump();
  __p(b.n);
}

void main() {
  __vybeMain();
  __check('1');
}
