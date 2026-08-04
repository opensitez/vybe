// vybe-test: dart/null_operators/null_aware_cascade_mutates_when_receiver_present
// origin: languages/dart/tests/dart/test_null_operators.rs

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

class Counter { int n = 0; void inc() { n++; } }
void __vybeMain() {
  Counter? c = Counter();
  c?.inc();
  __p(c?.n);
}

void main() {
  __vybeMain();
  __check('1');
}
