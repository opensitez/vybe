// vybe-test: dart/no_such_method/no_such_method_with_string_return_from_dynamic
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

class StrProxy {
  @override
  dynamic noSuchMethod(Invocation inv) {
    return 'proxy-${inv.positionalArguments.length}';
  }
}
void __vybeMain() {
  dynamic s = StrProxy();
  __p(s.msg('a', 'b'));
}

void main() {
  __vybeMain();
  __check('proxy-2');
}
