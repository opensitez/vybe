// vybe-test: dart/functions_core/top_level_function_returns_list_literal
// origin: languages/dart/tests/dart/test_functions_core.rs

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

List<int> makeRange(int n) {
  return [for (var i = 0; i < n; i++) i];
}
void __vybeMain() {
  __p(makeRange(4).join(','));
}

void main() {
  __vybeMain();
  __check('0,1,2,3');
}
