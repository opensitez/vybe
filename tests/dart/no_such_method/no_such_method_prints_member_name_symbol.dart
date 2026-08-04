// vybe-test: dart/no_such_method/no_such_method_prints_member_name_symbol
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

class Logger {
  @override
  dynamic noSuchMethod(Invocation inv) {
    __p(inv.memberName.toString());
    return 0;
  }
}
void __vybeMain() {
  dynamic l = Logger();
  l.fetchData();
}

void main() {
  __vybeMain();
  __check('Symbol("fetchData")');
}
