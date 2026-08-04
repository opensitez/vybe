// vybe-test: dart/enum_enhanced/enhanced_enum_field_large
// origin: languages/dart/tests/dart/test_enum_enhanced.rs

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

enum Big {
  mega(1000000);
  final int size;
  const Big(this.size);
}
void __vybeMain() {
  __p(Big.mega.size);
}

void main() {
  __vybeMain();
  __check('1000000');
}
