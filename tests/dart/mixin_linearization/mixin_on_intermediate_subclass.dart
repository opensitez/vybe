// vybe-test: dart/mixin_linearization/mixin_on_intermediate_subclass
// origin: languages/dart/tests/dart/test_mixin_linearization.rs

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

class Root {
  int depth = 1;
}
class Mid extends Root {}
mixin Tag on Mid {
  int tagged() {
    return depth + 10;
  }
}
class Leaf extends Mid with Tag {}
void __vybeMain() {
  __p(Leaf().tagged());
}

void main() {
  __vybeMain();
  __check('11');
}
