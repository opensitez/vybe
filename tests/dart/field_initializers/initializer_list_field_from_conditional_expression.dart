// vybe-test: dart/field_initializers/initializer_list_field_from_conditional_expression
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

class Pick {
  int chosen;
  Pick(bool useA, int a, int b) : chosen = useA ? a : b;
}
void __vybeMain() {
  __p(Pick(false, 1, 9).chosen);
}

void main() {
  __vybeMain();
  __check('9');
}
