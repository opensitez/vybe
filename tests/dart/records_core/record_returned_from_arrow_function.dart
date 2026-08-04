// vybe-test: dart/records_core/record_returned_from_arrow_function
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

({int x, int y}) point() => (x: 4, y: 8);
void __vybeMain() {
  var p = point();
  __p(p.x + p.y);
}

void main() {
  __vybeMain();
  __check('12');
}
