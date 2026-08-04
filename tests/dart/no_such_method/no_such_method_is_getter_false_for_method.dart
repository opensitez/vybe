// vybe-test: dart/no_such_method/no_such_method_is_getter_false_for_method
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

class Check {
  @override
  dynamic noSuchMethod(Invocation inv) {
    __p(inv.isGetter);
    return 0;
  }
}
void __vybeMain() {
  dynamic c = Check();
  c.run();
}

void main() {
  __vybeMain();
  __check('false');
}
