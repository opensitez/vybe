// vybe-test: dart/mixin_linearization/three_mixins_last_declared_wins_conflict
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

mixin M1 {
  int val() {
    return 1;
  }
}
mixin M2 {
  int val() {
    return 2;
  }
}
mixin M3 {
  int val() {
    return 3;
  }
}
class Z with M1, M2, M3 {}
void __vybeMain() {
  __p(Z().val());
}

void main() {
  __vybeMain();
  __check('3');
}
