// vybe-test: dart/typedefs_core/generic_typedef_predicate_filters_values
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

typedef Predicate<T> = bool Function(T);
List<T> keepIf<T>(List<T> items, Predicate<T> pred) {
  var out = <T>[];
  for (var item in items) {
    if (pred(item)) {
      out.add(item);
    }
  }
  return out;
}
void __vybeMain() {
  var evens = keepIf([1, 2, 3, 4], (n) => n % 2 == 0);
  __p(evens.join(','));
}

void main() {
  __vybeMain();
  __check('2,4');
}
