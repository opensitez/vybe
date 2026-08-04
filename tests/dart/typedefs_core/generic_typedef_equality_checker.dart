// vybe-test: dart/typedefs_core/generic_typedef_equality_checker
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

typedef Eq<T> = bool Function(T, T);
bool allEqual<T>(List<T> items, Eq<T> same) {
  if (items.isEmpty) {
    return true;
  }
  var first = items.first;
  for (var item in items) {
    if (!same(item, first)) {
      return false;
    }
  }
  return true;
}
void __vybeMain() {
  __p(allEqual(['x', 'x', 'x'], (a, b) => a == b));
  __p(allEqual(['x', 'y'], (a, b) => a == b));
}

void main() {
  __vybeMain();
  __check('true\nfalse');
}
