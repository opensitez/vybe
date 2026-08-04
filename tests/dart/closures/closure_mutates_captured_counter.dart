// vybe-test: dart/closures/closure_mutates_captured_counter
// origin: languages/dart/tests/dart/test_closures.rs

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

void __vybeMain() {
  var count = 0;
  var inc = () {
    count++;
  };
  inc();
  inc();
  inc();
  __p(count);
}

void main() {
  __vybeMain();
  __check('3');
}
