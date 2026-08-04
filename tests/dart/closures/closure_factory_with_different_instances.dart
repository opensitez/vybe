// vybe-test: dart/closures/closure_factory_with_different_instances
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

int Function(int) makeOffset(int n) {
  return (x) => x - n;
}
void __vybeMain() {
  var sub3 = makeOffset(3);
  var sub5 = makeOffset(5);
  __p(sub3(10));
  __p(sub5(10));
}

void main() {
  __vybeMain();
  __check('7\n5');
}
