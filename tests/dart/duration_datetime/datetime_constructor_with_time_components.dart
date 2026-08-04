// vybe-test: dart/duration_datetime/datetime_constructor_with_time_components
// origin: languages/dart/tests/dart/test_duration_datetime.rs

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

void __vybeMain() {
  var dt = DateTime(2024, 1, 1, 14, 30, 45);
  __p(dt.hour);
  __p(dt.minute);
  __p(dt.second);
}

void main() {
  __vybeMain();
  __check('14\n30\n45');
}
