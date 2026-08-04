// vybe-test: dart/class_modifiers/interface_and_final_independent_types
// origin: languages/dart/tests/dart/test_class_modifiers.rs

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

interface class Port {
  int number();
}
final class Endpoint implements Port {
  int _n;
  Endpoint(this._n);
  @override
  int number() {
    return _n;
  }
}
void __vybeMain() {
  __p(Endpoint(8080).number());
}

void main() {
  __vybeMain();
  __check('8080');
}
