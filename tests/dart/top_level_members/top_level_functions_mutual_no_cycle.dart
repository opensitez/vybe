// vybe-test: dart/top_level_members/top_level_functions_mutual_no_cycle
// origin: languages/dart/tests/dart/test_top_level_members.rs

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

int toA(int n) {
  return n + 10;
}
int toB(int n) {
  return toA(n) + 5;
}
void __vybeMain() {
  __p(toB(1));
}

void main() {
  __vybeMain();
  __check('16');
}
