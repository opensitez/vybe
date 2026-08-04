// vybe-test: dart/mixins_core/mixin_method_uses_this
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

mixin Self {
  int id() {
    return 1;
  }
  int same() {
    return this.id();
  }
}
class Host with Self {}
void __vybeMain() {
  __p(Host().same());
}

void main() {
  __vybeMain();
  __check('1');
}
