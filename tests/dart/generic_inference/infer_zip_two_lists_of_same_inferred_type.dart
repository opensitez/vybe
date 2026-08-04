// vybe-test: dart/generic_inference/infer_zip_two_lists_of_same_inferred_type
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

List<Pair<T, T>> zipSame<T>(List<T> a, List<T> b) {
  var out = <Pair<T, T>>[];
  for (var i = 0; i < a.length && i < b.length; i++) {
    out.add(Pair(a[i], b[i]));
  }
  return out;
}
class Pair<A, B> {
  A first;
  B second;
  Pair(this.first, this.second);
}
void __vybeMain() {
  var z = zipSame([1, 2], [3, 4]);
  __p(z[0].first);
  __p(z[1].second);
}

void main() {
  __vybeMain();
  __check('1\n4');
}
