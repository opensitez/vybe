// vybe-test: dart/no_such_method/no_such_method_with_multiple_named_args
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

class Config {
  @override
  dynamic noSuchMethod(Invocation inv) {
    var a = inv.namedArguments[#a] as int;
    var b = inv.namedArguments[#b] as int;
    return a + b;
  }
}
void __vybeMain() {
  dynamic c = Config();
  __p(c.merge(a: 3, b: 4));
}

void main() {
  __vybeMain();
  __check('7');
}
