// vybe-test: dart/records_core/record_stored_in_list
// origin: languages/dart/tests/dart/test_records_core.rs

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
  var points = [(0, 0), (1, 2), (3, 4)];
  __p(points.length);
  __p(points[1].$1);
  __p(points[1].$2);
}

void main() {
  __vybeMain();
  __check('3\n1\n2');
}
