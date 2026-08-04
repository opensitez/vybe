// vybe-test: dart/no_such_method/dynamic_call_after_assignment
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

class Box {
  @override
  dynamic noSuchMethod(Invocation inv) {
    return 100;
  }
}
void __vybeMain() {
  dynamic b;
  b = Box();
  __p(b.fetch());
}

void main() {
  __vybeMain();
  __check('100');
}
