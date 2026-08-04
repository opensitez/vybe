// vybe-test: dart/getters_setters/getter_derived_from_other_getter
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

class Circle {
  double _radius = 2.0;
  double get radius {
    return _radius;
  }
  double get diameter {
    return radius * 2;
  }
}
void __vybeMain() {
  __p(Circle().diameter);
}

void main() {
  __vybeMain();
  __check('4.0');
}
