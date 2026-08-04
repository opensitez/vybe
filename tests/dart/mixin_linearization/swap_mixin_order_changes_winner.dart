// vybe-test: dart/mixin_linearization/swap_mixin_order_changes_winner
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

mixin First {
  int pick() {
    return 100;
  }
}
mixin Second {
  int pick() {
    return 200;
  }
}
class One with First, Second {}
class Two with Second, First {}
void __vybeMain() {
  __p(One().pick() + Two().pick());
}

void main() {
  __vybeMain();
  __check('300');
}
