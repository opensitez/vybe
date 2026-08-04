// vybe-test: dart/typedefs_core/generic_typedef_chain_mapper_and_predicate
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

typedef Mapper<T, R> = R Function(T);
typedef Predicate<R> = bool Function(R);
bool anyMatch<T, R>(List<T> items, Mapper<T, R> map, Predicate<R> pred) {
  for (var item in items) {
    if (pred(map(item))) {
      return true;
    }
  }
  return false;
}
void __vybeMain() {
  __p(anyMatch(['aa', 'b', 'ccc'], (s) => s.length, (len) => len > 2));
  __p(anyMatch(['a', 'bb'], (s) => s.length, (len) => len > 3));
}

void main() {
  __vybeMain();
  __check('true\nfalse');
}
