// vybe-test: dart/no_such_method/no_such_method_positional_args_types
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

class Types {
  @override
  dynamic noSuchMethod(Invocation inv) {
    var a = inv.positionalArguments[0];
    var b = inv.positionalArguments[1];
    return '$a:$b';
  }
}
void __vybeMain() {
  dynamic t = Types();
  __p(t.pair(1, 'x'));
}

void main() {
  __vybeMain();
  __check('1:x');
}
