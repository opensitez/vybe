// vybe-test: dart/mixin_linearization/class_method_overridden_by_rightmost_mixin
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

class Host {
  int score() {
    return 1;
  }
}
mixin Boost {
  int score() {
    return 10;
  }
}
class Player extends Host with Boost {}
void __vybeMain() {
  __p(Player().score());
}

void main() {
  __vybeMain();
  __check('10');
}
