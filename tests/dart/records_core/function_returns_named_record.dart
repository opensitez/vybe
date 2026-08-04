// vybe-test: dart/records_core/function_returns_named_record
// origin: languages/dart/tests/dart/test_records_core.rs

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

({String name, int score}) topPlayer() {
  return (name: 'Zara', score: 100);
}
void __vybeMain() {
  var p = topPlayer();
  __p(p.name);
  __p(p.score);
}

void main() {
  __vybeMain();
  __check('Zara\n100');
}
