// vybe-test: dart/enum_enhanced/enhanced_enum_method_in_values_loop
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

enum Digit {
  d0(0),
  d1(1),
  d2(2);
  final int num;
  const Digit(this.num);
}
void __vybeMain() {
  var sum = 0;
  for (var d in Digit.values) {
    sum += d.num;
  }
  __p(sum);
}

void main() {
  __vybeMain();
  __check('3');
}
