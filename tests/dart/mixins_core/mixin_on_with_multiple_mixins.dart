// vybe-test: dart/mixins_core/mixin_on_with_multiple_mixins
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

class Engine {
  int power = 100;
}
mixin Turbo on Engine {
  int boost() {
    return power + 50;
  }
}
mixin Eco on Engine {
  int save() {
    return power - 20;
  }
}
class Motor extends Engine with Turbo, Eco {}
void __vybeMain() {
  var m = Motor();
  __p(m.boost());
}

void main() {
  __vybeMain();
  __check('150');
}
