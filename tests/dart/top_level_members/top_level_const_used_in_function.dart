// vybe-test: dart/top_level_members/top_level_const_used_in_function
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

const factor = 5;
int scale(int n) {
  return n * factor;
}
void __vybeMain() {
  __p(scale(6));
}

void main() {
  __vybeMain();
  __check('30');
}
