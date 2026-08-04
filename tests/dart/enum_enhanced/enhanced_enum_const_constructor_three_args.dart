// vybe-test: dart/enum_enhanced/enhanced_enum_const_constructor_three_args
// origin: languages/dart/tests/dart/test_enum_enhanced.rs

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

enum Planet {
  earth(5.97e24, 6371, 1),
  mars(6.39e23, 3389, 2);
  final double mass;
  final double radius;
  final int order;
  const Planet(this.mass, this.radius, this.order);
}
void __vybeMain() {
  __p(Planet.mars.order);
}

void main() {
  __vybeMain();
  __check('2');
}
