// vybe-test: dart/type_casts/runtime_type_of_custom_class
// origin: languages/dart/tests/dart/test_type_casts.rs

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

class Point { int x = 0; int y = 0; }
void __vybeMain() {
  var p = Point();
  __p(p.runtimeType == Point);
}

void main() {
  __vybeMain();
  __check('true');
}
