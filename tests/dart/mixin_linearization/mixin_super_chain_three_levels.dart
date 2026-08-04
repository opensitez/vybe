// vybe-test: dart/mixin_linearization/mixin_super_chain_three_levels
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

mixin A {
  String chain() {
    return 'A';
  }
}
mixin B on Object {
  String chain() {
    return super.chain() + 'B';
  }
}
mixin C on Object {
  String chain() {
    return super.chain() + 'C';
  }
}
class D with A, B, C {}
void __vybeMain() {
  __p(D().chain());
}

void main() {
  __vybeMain();
  __check('ABC');
}
