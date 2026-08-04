// vybe-test: dart/generics_core/top_level_generic_swap_pair
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

Pair<T, U> swap<T, U>(Pair<T, U> p) {
  return Pair(p.second, p.first);
}
class Pair<T, U> {
  T first;
  U second;
  Pair(this.first, this.second);
}
void __vybeMain() {
  var p = Pair(1, 'a');
  var s = swap(p);
  __p(s.first);
  __p(s.second);
}

void main() {
  __vybeMain();
  __check('a\n1');
}
