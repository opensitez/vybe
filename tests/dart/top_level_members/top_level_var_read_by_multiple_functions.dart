// vybe-test: dart/top_level_members/top_level_var_read_by_multiple_functions
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

int shared = 5;
int readShared() {
  return shared;
}
int bumpShared() {
  shared = shared + 1;
  return shared;
}
void __vybeMain() {
  __p(readShared());
  __p(bumpShared());
  __p(readShared());
}

void main() {
  __vybeMain();
  __check('5\n6\n6');
}
