// vybe-test: dart/factory_constructors_deep/factory_redirect_named_to_generative_named
// origin: languages/dart/tests/dart/test_factory_constructors_deep.rs

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

class Pair {
  int a;
  int b;
  Pair(this.a, this.b);
  Pair.same(int v) : a = v, b = v;
  factory Pair.fromSame(int v) => Pair.same(v);
}
void __vybeMain() {
  var p = Pair.fromSame(3);
  __p(p.a + p.b);
}

void main() {
  __vybeMain();
  __check('6');
}
