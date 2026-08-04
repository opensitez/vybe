// vybe-test: dart/generic_inference/infer_fold_string_concat_from_seed
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

T foldList<T>(List<T> items, T Function(T, T) combine) {
  var acc = items.first;
  for (var i = 1; i < items.length; i++) {
    acc = combine(acc, items[i]);
  }
  return acc;
}
void __vybeMain() {
  __p(foldList(['a', 'b'], (a, b) => a + b));
}

void main() {
  __vybeMain();
  __check('ab');
}
