// vybe-test: dart/const_final/late_final_result
// origin: languages/dart/tests/dart/test_const_final.rs

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
  late final int start;
  Counter(int v) { start = v; }
}
void __vybeMain() {
  var c = Counter(10);
  __p(c.start);
}

void main() {
  __vybeMain();
  __check('10');
}
