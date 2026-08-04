// vybe-test: dart/cascades/cascade_custom_counter_increment_chain_accumulates
// origin: languages/dart/tests/dart/test_cascades.rs

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

class Counter {
  int value = 0;
  void bump() { value += 1; }
}
void __vybeMain() {
  var tally = Counter();
  tally..bump()..bump()..bump();
  __p(tally.value);
}

void main() {
  __vybeMain();
  __check('3');
}
