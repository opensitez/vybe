// vybe-test: dart/field_initializers/three_field_initializer_list_left_to_right
// origin: languages/dart/tests/dart/test_field_initializers.rs

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

class Triple {
  int a;
  int b;
  int c;
  Triple(int x) : a = x, b = x + 1, c = x + 2;
}
void __vybeMain() {
  var t = Triple(10);
  __p(t.a + t.b + t.c);
}

void main() {
  __vybeMain();
  __check('33');
}
