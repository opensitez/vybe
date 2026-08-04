// vybe-test: dart/mixin_linearization/mixin_conflict_resolution_with_extends_and_with
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
  String who() {
    return 'root';
  }
}
mixin Layer {
  String who() {
    return 'layer';
  }
}
class Leaf extends Root with Layer {}
void __vybeMain() {
  __p(Leaf().who());
}

void main() {
  __vybeMain();
  __check('layer');
}
