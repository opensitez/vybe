// vybe-test: dart/no_such_method/no_such_method_zero_arg_method
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

class Zero {
  @override
  dynamic noSuchMethod(Invocation inv) {
    return inv.positionalArguments.isEmpty;
  }
}
void __vybeMain() {
  dynamic z = Zero();
  __p(z.ping());
}

void main() {
  __vybeMain();
  __check('true');
}
