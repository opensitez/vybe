// vybe-test: dart/enum_enhanced/enhanced_enum_switch_five_members
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

enum Weekday { mon, tue, wed, thu, fri }
int num(Weekday w) {
  switch (w) {
    case Weekday.mon:
      return 1;
    case Weekday.tue:
      return 2;
    case Weekday.wed:
      return 3;
    case Weekday.thu:
      return 4;
    case Weekday.fri:
      return 5;
  }
}
void __vybeMain() {
  __p(num(Weekday.wed));
}

void main() {
  __vybeMain();
  __check('3');
}
