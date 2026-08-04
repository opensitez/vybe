// vybe-test: dart/field_initializers/initializer_list_division_truncates
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

class Ratio {
  int whole;
  Ratio(int a, int b) : whole = a ~/ b;
}
void __vybeMain() {
  __p(Ratio(7, 2).whole);
}

void main() {
  __vybeMain();
  __check('3');
}
