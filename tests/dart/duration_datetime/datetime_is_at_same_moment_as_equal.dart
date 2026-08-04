// vybe-test: dart/duration_datetime/datetime_is_at_same_moment_as_equal
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
  var a = DateTime(2024, 5, 17, 8, 0, 0);
  var b = DateTime(2024, 5, 17, 8, 0, 0);
  __p(a.isAtSameMomentAs(b));
}

void main() {
  __vybeMain();
  __check('true');
}
