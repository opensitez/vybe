// vybe-test: dart/typedefs_core/typedef_stored_in_list_and_invoked
// origin: languages/dart/tests/dart/test_typedefs_core.rs

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

typedef Step = int Function(int);
int addOne(int n) {
  return n + 1;
}
int addTwo(int n) {
  return n + 2;
}
void __vybeMain() {
  List<Step> steps = [addOne, addTwo];
  __p(steps[0](5));
  __p(steps[1](5));
}

void main() {
  __vybeMain();
  __check('6\n7');
}
