// vybe-test: dart/static_members/static_method_with_single_arg
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

class Square {
  static int sq(int n) {
    return n * n;
  }
}
void __vybeMain() {
  __p(Square.sq(6));
}

void main() {
  __vybeMain();
  __check('36');
}
