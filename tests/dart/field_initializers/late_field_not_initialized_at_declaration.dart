// vybe-test: dart/field_initializers/late_field_not_initialized_at_declaration
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

class Lazy {
  late int value;
  Lazy(int v) : value = v;
}
void __vybeMain() {
  __p(Lazy(9).value);
}

void main() {
  __vybeMain();
  __check('9');
}
