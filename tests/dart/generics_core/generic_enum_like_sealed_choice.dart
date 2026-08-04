// vybe-test: dart/generics_core/generic_enum_like_sealed_choice
// origin: languages/dart/tests/dart/test_generics_core.rs

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

class Choice<T> {
  T? some;
  bool get isSome {
    return some != null;
  }
  Choice.some(this.some);
  Choice.none() : some = null;
}
void __vybeMain() {
  var c = Choice<int>.some(9);
  __p(c.isSome);
  __p(c.some);
}

void main() {
  __vybeMain();
  __check('true\n9');
}
