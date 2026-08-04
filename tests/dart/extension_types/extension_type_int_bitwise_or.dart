// vybe-test: dart/extension_types/extension_type_int_bitwise_or
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

extension type Flags(int bits) {
  Flags or(Flags other) {
    return Flags(bits | other.bits);
  }
}
void __vybeMain() {
  Flags a = Flags(1);
  Flags b = Flags(2);
  __p(a.or(b).bits);
}

void main() {
  __vybeMain();
  __check('3');
}
