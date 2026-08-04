// vybe-test: dart/duration_datetime/datetime_difference_in_days
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
  var start = DateTime(2024, 1, 1);
  var end = DateTime(2024, 1, 11);
  __p(end.difference(start).inDays);
}

void main() {
  __vybeMain();
  __check('10.0');
}
