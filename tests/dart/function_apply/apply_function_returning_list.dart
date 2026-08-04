// vybe-test: dart/function_apply/apply_function_returning_list
// origin: languages/dart/tests/dart/test_function_apply.rs

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

List<int> range(int start, int end) {
  return [start, end];
}
void __vybeMain() {
  var list = Function.apply(range, [1, 3]) as List;
  __p(list.length);
  __p(list[0]);
  __p(list[1]);
}

void main() {
  __vybeMain();
  __check('2\n1\n3');
}
