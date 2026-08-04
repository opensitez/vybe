// vybe-test: dart/extensions/list_getter_first_element_value
// origin: languages/dart/tests/dart/test_extensions.rs

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

extension ListHead on List<int> {
  int get firstVal => this[0];
}
void __vybeMain() {
  __p([9, 8, 7].firstVal);
}

void main() {
  __vybeMain();
  __check('9');
}
