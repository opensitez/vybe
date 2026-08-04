// vybe-test: dart/named_parameters/two_required_named_both_provided
// origin: languages/dart/tests/dart/test_named_parameters.rs

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

void move({required int dx, required int dy}) {
  __p('$dx,$dy');
}
void __vybeMain() {
  move(dx: 3, dy: -1);
}

void main() {
  __vybeMain();
  __check('3,-1');
}
