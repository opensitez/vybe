// vybe-test: dart/getters_setters/computed_getter_multiplies_two_fields
// origin: languages/dart/tests/dart/test_getters_setters.rs

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

class Rect {
  int width = 3;
  int height = 4;
  int get area {
    return width * height;
  }
}
void __vybeMain() {
  __p(Rect().area);
}

void main() {
  __vybeMain();
  __check('12');
}
