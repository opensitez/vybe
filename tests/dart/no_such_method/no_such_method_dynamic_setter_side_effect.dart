// vybe-test: dart/no_such_method/no_such_method_dynamic_setter_side_effect
// origin: languages/dart/tests/dart/test_no_such_method.rs

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

class Store {
  int stored = 0;
  @override
  dynamic noSuchMethod(Invocation inv) {
    if (inv.isSetter) {
      stored = inv.positionalArguments[0] as int;
    }
    return null;
  }
}
void __vybeMain() {
  dynamic s = Store();
  s.value = 15;
  __p(s.stored);
}

void main() {
  __vybeMain();
  __check('15');
}
