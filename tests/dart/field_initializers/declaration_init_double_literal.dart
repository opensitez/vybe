// vybe-test: dart/field_initializers/declaration_init_double_literal
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

class Measure {
  double pi = 3.14;
}
void __vybeMain() {
  __p(Measure().pi > 3.0);
}

void main() {
  __vybeMain();
  __check('true');
}
