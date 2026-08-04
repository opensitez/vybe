// vybe-test: dart/field_initializers/initializer_list_multiple_asserts_with_fields
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

class Range {
  int lo;
  int hi;
  Range(int a, int b) : lo = a, hi = b, assert(a <= b);
}
void __vybeMain() {
  __p(Range(2, 8).hi);
}

void main() {
  __vybeMain();
  __check('8');
}
