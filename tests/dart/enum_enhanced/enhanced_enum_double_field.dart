// vybe-test: dart/enum_enhanced/enhanced_enum_double_field
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

enum Rate {
  half(0.5),
  full(1.0);
  final double factor;
  const Rate(this.factor);
}
void __vybeMain() {
  __p(Rate.half.factor + Rate.full.factor);
}

void main() {
  __vybeMain();
  __check('1.5');
}
