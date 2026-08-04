// vybe-test: dart/duration_datetime/datetime_add_then_difference_roundtrip
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
  var base = DateTime(2024, 7, 1);
  var span = Duration(days: 14);
  var target = base.add(span);
  __p(target.difference(base).inDays);
}

void main() {
  __vybeMain();
  __check('14.0');
}
