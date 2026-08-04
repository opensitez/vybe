// vybe-test: dart/generic_inference/infer_map_values_from_nested_list_literal
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

Map<int, List<String>> build() {
  return {1: ['a'], 2: ['b', 'c']};
}
void __vybeMain() {
  var m = build();
  __p(m[2]!.length);
  __p(m[2]![0]);
}

void main() {
  __vybeMain();
  __check('2\nb');
}
