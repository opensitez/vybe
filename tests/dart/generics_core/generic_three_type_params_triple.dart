// vybe-test: dart/generics_core/generic_three_type_params_triple
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

class Triple<A, B, C> {
  A first;
  B second;
  C third;
  Triple(this.first, this.second, this.third);
}
void __vybeMain() {
  var t = Triple(1, 'x', true);
  __p(t.first);
  __p(t.second);
  __p(t.third);
}

void main() {
  __vybeMain();
  __check('1\nx\ntrue');
}
