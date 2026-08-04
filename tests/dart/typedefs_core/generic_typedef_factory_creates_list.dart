// vybe-test: dart/typedefs_core/generic_typedef_factory_creates_list
// origin: languages/dart/tests/dart/test_typedefs_core.rs

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

typedef ListFactory<T> = List<T> Function();
List<int> makeInts() {
  return [1, 2, 3];
}
void __vybeMain() {
  ListFactory<int> factory = makeInts;
  __p(factory().join(','));
}

void main() {
  __vybeMain();
  __check('1,2,3');
}
