// vybe-test: dart/extension_types/extension_type_int_equality_different_value
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

extension type Id(int value) {
  bool sameAs(Id other) {
    return value == other.value;
  }
}
void __vybeMain() {
  Id a = Id(5);
  Id b = Id(6);
  __p(a.sameAs(b));
}

void main() {
  __vybeMain();
  __check('false');
}
