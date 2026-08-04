// vybe-test: dart/extension_types/extension_type_int_modulo_representation
// origin: languages/dart/tests/dart/test_extension_types.rs

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

extension type Quantity(int amount) {
  int remainder(int divisor) {
    return amount % divisor;
  }
}
void __vybeMain() {
  Quantity q = Quantity(17);
  __p(q.remainder(5));
}

void main() {
  __vybeMain();
  __check('2');
}
