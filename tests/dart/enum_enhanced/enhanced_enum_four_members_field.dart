// vybe-test: dart/enum_enhanced/enhanced_enum_four_members_field
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

enum Season {
  spring(1),
  summer(2),
  autumn(3),
  winter(4);
  final int month;
  const Season(this.month);
}
void __vybeMain() {
  __p(Season.winter.month);
}

void main() {
  __vybeMain();
  __check('4');
}
