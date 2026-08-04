// vybe-test: dart/abstract_members/abstract_getter_setter_pair
// origin: languages/dart/tests/dart/test_abstract_members.rs

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

abstract class Temperature {
  double get celsius;
  set celsius(double v);
}
class Room extends Temperature {
  double _c = 20.0;
  double get celsius => _c;
  set celsius(double v) {
    _c = v;
  }
}
void __vybeMain() {
  var r = Room();
  r.celsius = 25.0;
  __p(r.celsius);
}

void main() {
  __vybeMain();
  __check('25.0');
}
