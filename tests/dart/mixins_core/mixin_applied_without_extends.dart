// vybe-test: dart/mixins_core/mixin_applied_without_extends
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

mixin M {
  int get flag {
    return 1;
  }
}
class Plain with M {}
void __vybeMain() {
  __p(Plain().flag);
}

void main() {
  __vybeMain();
  __check('1');
}
