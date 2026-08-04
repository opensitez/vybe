// vybe-test: dart/no_such_method/no_such_method_dynamic_method_returns_list_element
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

class ListProxy {
  final List<int> data = [10, 20, 30];
  @override
  dynamic noSuchMethod(Invocation inv) {
    if (inv.isMethod && inv.memberName == #at) {
      return data[inv.positionalArguments[0] as int];
    }
    return null;
  }
}
void __vybeMain() {
  dynamic p = ListProxy();
  __p(p.at(1));
}

void main() {
  __vybeMain();
  __check('20');
}
