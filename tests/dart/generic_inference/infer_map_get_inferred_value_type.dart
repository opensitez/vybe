// vybe-test: dart/generic_inference/infer_map_get_inferred_value_type
// origin: languages/dart/tests/dart/test_generic_inference.rs

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

V? lookup<K, V>(Map<K, V> map, K key) {
  return map[key];
}
void __vybeMain() {
  var table = {'a': 1, 'b': 2};
  __p(lookup(table, 'b'));
}

void main() {
  __vybeMain();
  __check('2');
}
