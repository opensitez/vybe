// vybe-test: dart/typedefs_core/generic_typedef_mapper_transforms_list
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
List<R> mapAll<T, R>(List<T> items, Mapper<T, R> fn) {
  var out = <R>[];
  for (var item in items) {
    out.add(fn(item));
  }
  return out;
}
void __vybeMain() {
  var lengths = mapAll(['a', 'bb'], (String s) => s.length);
  __p(lengths.join(','));
}

void main() {
  __vybeMain();
  __check('1,2');
}
