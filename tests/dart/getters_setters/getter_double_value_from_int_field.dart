// vybe-test: dart/getters_setters/getter_double_value_from_int_field
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

class Scaler {
  int _factor = 3;
  int get factor {
    return _factor;
  }
  double get scaled {
    return _factor * 1.5;
  }
}
void __vybeMain() {
  __p(Scaler().scaled);
}

void main() {
  __vybeMain();
  __check('4.5');
}
