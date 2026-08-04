// vybe-test: dart/mixin_linearization/diamond_like_mixin_resolution_last_wins
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

mixin X {
  String id() {
    return 'x';
  }
}
mixin Y {
  String id() {
    return 'y';
  }
}
class Combo with X, Y {}
void __vybeMain() {
  __p(Combo().id());
}

void main() {
  __vybeMain();
  __check('y');
}
