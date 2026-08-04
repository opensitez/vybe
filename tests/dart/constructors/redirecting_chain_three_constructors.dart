// vybe-test: dart/constructors/redirecting_chain_three_constructors
// origin: languages/dart/tests/dart/test_constructors.rs

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

class N {
  int v;
  N(this.v);
  N.one() : this(1);
  N.copy() : this.one();
}
void __vybeMain() {
  __p(N.copy().v);
}

void main() {
  __vybeMain();
  __check('1');
}
