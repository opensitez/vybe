// vybe-test: dart/classes_advanced/to_string_result
// origin: languages/dart/tests/dart/test_classes_advanced.rs

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

class Point { int x; int y; Point(this.x, this.y); String toString() => '($x, $y)'; }
void __vybeMain() { var p = Point(3, 4); __p(p); }

void main() {
  __vybeMain();
  __check('(3, 4)');
}
