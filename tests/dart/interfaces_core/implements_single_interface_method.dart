// vybe-test: dart/interfaces_core/implements_single_interface_method
// origin: languages/dart/tests/dart/test_interfaces_core.rs

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

abstract class Drawable {
  String draw();
}
class Circle implements Drawable {
  String draw() {
    return 'circle';
  }
}
void __vybeMain() {
  __p(Circle().draw());
}

void main() {
  __vybeMain();
  __check('circle');
}
