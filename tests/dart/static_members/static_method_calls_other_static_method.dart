// vybe-test: dart/static_members/static_method_calls_other_static_method
// origin: languages/dart/tests/dart/test_static_members.rs

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

class Chain {
  static int a(int n) {
    return n + 1;
  }
  static int b(int n) {
    return a(n) * 2;
  }
}
void __vybeMain() {
  __p(Chain.b(3));
}

void main() {
  __vybeMain();
  __check('8');
}
