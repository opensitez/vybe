// vybe-test: dart/enum_enhanced/enhanced_enum_switch_with_field_access
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

enum Level {
  low(1),
  high(10);
  final int power;
  const Level(this.power);
}
int boost(Level l) {
  switch (l) {
    case Level.low:
      return l.power;
    case Level.high:
      return l.power * 2;
  }
}
void __vybeMain() {
  __p(boost(Level.high));
}

void main() {
  __vybeMain();
  __check('20');
}
