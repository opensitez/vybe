// vybe-test: dart/enums_core/enhanced_enum_getter_false_on_weekday
// origin: languages/dart/tests/dart/test_enums_core.rs

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

enum Day { monday, tuesday, wednesday, thursday, friday, saturday, sunday;
  bool get isWeekend => this == Day.saturday || this == Day.sunday;
}
void __vybeMain() {
  __p(Day.monday.isWeekend);
}

void main() {
  __vybeMain();
  __check('false');
}
