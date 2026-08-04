// vybe-test: dart/mixin_linearization/five_mixins_rightmost_tag_wins
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

mixin T1 {
  String tag() {
    return '1';
  }
}
mixin T2 {
  String tag() {
    return '2';
  }
}
mixin T3 {
  String tag() {
    return '3';
  }
}
mixin T4 {
  String tag() {
    return '4';
  }
}
mixin T5 {
  String tag() {
    return '5';
  }
}
class All with T1, T2, T3, T4, T5 {}
void __vybeMain() {
  __p(All().tag());
}

void main() {
  __vybeMain();
  __check('5');
}
