// vybe-test: dart/generics_core/generic_pair_equality_by_fields
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

class Pair<T> {
  T a;
  T b;
  Pair(this.a, this.b);
  bool sameFirst(Pair<T> other) {
    return a == other.a;
  }
}
void __vybeMain() {
  var p1 = Pair(1, 2);
  var p2 = Pair(1, 9);
  __p(p1.sameFirst(p2));
}

void main() {
  __vybeMain();
  __check('true');
}
