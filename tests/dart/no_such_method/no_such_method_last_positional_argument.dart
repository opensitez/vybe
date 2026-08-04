// vybe-test: dart/no_such_method/no_such_method_last_positional_argument
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

class Last {
  @override
  dynamic noSuchMethod(Invocation inv) {
    return inv.positionalArguments.last;
  }
}
void __vybeMain() {
  dynamic l = Last();
  __p(l.pick(1, 2, 9));
}

void main() {
  __vybeMain();
  __check('9');
}
