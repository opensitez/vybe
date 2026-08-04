// vybe-test: dart/classes_core/to_string_override_formats_instance
// origin: languages/dart/tests/dart/test_classes_core.rs

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

class Point {
  int x = 1;
  int y = 2;
  String toString() {
    return '($x,$y)';
  }
}
void __vybeMain() {
  __p(Point());
}

void main() {
  __vybeMain();
  __check('(1,2)');
}
