// vybe-test: dart/const_final/final_computed
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

void __vybeMain() { final a = 6; final b = 7; final c = a * b; __p(c); }

void main() {
  __vybeMain();
  __check('42');
}
