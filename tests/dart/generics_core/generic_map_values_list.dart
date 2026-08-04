// vybe-test: dart/generics_core/generic_map_values_list
// origin: languages/dart/tests/dart/test_generics_core.rs

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

Map<K, List<V>> bucket<K, V>(List<V> items, V Function(V) keyFn) {
  var map = <K, List<V>>{};
  for (var item in items) {
    var key = keyFn(item);
    map.putIfAbsent(key, () => []).add(item);
  }
  return map;
}
void __vybeMain() {
  var m = bucket([1, 2, 3, 4], (n) => n % 2);
  __p(m[0]!.length);
  __p(m[1]!.length);
}

void main() {
  __vybeMain();
  __check('2\n2');
}
