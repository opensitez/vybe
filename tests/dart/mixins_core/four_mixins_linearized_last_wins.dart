// vybe-test: dart/mixins_core/four_mixins_linearized_last_wins
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

mixin A {
  String tag() {
    return 'a';
  }
}
mixin B {
  String tag() {
    return 'b';
  }
}
mixin C {
  String tag() {
    return 'c';
  }
}
mixin D {
  String tag() {
    return 'd';
  }
}
class X with A, B, C, D {}
void __vybeMain() {
  __p(X().tag());
}

void main() {
  __vybeMain();
  __check('d');
}
