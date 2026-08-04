// vybe-test: dart/constructors/redirecting_to_primary_with_different_args
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

class Range {
  int lo;
  int hi;
  Range(this.lo, this.hi);
  Range.single(int n) : this(n, n);
}
void __vybeMain() {
  var r = Range.single(4);
  __p(r.hi);
}

void main() {
  __vybeMain();
  __check('4');
}
