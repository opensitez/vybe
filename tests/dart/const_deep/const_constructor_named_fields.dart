// vybe-test: dart/const_deep/const_constructor_named_fields
// origin: languages/dart/tests/dart/test_const_deep.rs

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

class Size {
  final int width;
  final int height;
  const Size({required this.width, required this.height});
}
void __vybeMain() {
  const s = Size(width: 5, height: 6);
  __p(s.width * s.height);
}

void main() {
  __vybeMain();
  __check('30');
}
