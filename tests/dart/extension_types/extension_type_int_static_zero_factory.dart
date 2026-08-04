// vybe-test: dart/extension_types/extension_type_int_static_zero_factory
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

extension type Count(int value) {
  static Count zero() {
    return Count(0);
  }
}
void __vybeMain() {
  Count c = Count.zero();
  __p(c.value);
}

void main() {
  __vybeMain();
  __check('0');
}
