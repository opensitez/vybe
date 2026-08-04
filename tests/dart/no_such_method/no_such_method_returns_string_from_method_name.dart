// vybe-test: dart/no_such_method/no_such_method_returns_string_from_method_name
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

class NameEcho {
  @override
  dynamic noSuchMethod(Invocation inv) {
    var s = inv.memberName.toString();
    return s.replaceAll('Symbol(\"', '').replaceAll('\")', '');
  }
}
void __vybeMain() {
  dynamic n = NameEcho();
  __p(n.hello());
}

void main() {
  __vybeMain();
  __check('hello');
}
