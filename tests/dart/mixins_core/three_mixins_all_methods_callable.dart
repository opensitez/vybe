// vybe-test: dart/mixins_core/three_mixins_all_methods_callable
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

mixin M1 {
  int one() {
    return 1;
  }
}
mixin M2 {
  int two() {
    return 2;
  }
}
mixin M3 {
  int three() {
    return 3;
  }
}
class All with M1, M2, M3 {}
void __vybeMain() {
  var a = All();
  __p(a.one() + a.two() + a.three());
}

void main() {
  __vybeMain();
  __check('6');
}
