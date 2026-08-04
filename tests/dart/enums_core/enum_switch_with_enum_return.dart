// vybe-test: dart/enums_core/enum_switch_with_enum_return
// origin: languages/dart/tests/dart/test_enums_core.rs

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

enum Shape { circle, square }
String label(Shape s) {
  switch (s) {
    case Shape.circle:
      return 'round';
    default:
      return 'flat';
  }
}
void __vybeMain() {
  __p(label(Shape.square));
}

void main() {
  __vybeMain();
  __check('flat');
}
