// vybe-test: dart/generics/generic_pair_result
// origin: languages/dart/tests/dart/test_generics.rs

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

class Pair<A, B> { A first; B second; Pair(this.first, this.second); }
void __vybeMain() { var p = Pair(1, 'one'); __p(p.first); }

void main() {
  __vybeMain();
  __check('1');
}
