// vybe-test: dart/type_casts/as_cast_subclass_to_parent
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

class Shape { String kind = 'shape'; }
class Circle extends Shape { double r = 1.0; }
void __vybeMain() {
  Circle c = Circle();
  var s = c as Shape;
  __p(s.kind);
}

void main() {
  __vybeMain();
  __check('shape');
}
