// vybe-test: dart/mixins_core/later_mixin_wins_over_earlier_for_same_name
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

mixin First {
  int pick() {
    return 1;
  }
}
mixin Second {
  int pick() {
    return 2;
  }
}
class Combo with First, Second {}
void __vybeMain() {
  __p(Combo().pick());
}

void main() {
  __vybeMain();
  __check('2');
}
