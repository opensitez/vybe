// vybe-test: dart/mixin_linearization/reversed_pair_mixins_opposite_winners
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

mixin Red {
  int hue() {
    return 1;
  }
}
mixin Blue {
  int hue() {
    return 2;
  }
}
class RB with Red, Blue {}
class BR with Blue, Red {}
void __vybeMain() {
  __p(RB().hue() + BR().hue());
}

void main() {
  __vybeMain();
  __check('3');
}
