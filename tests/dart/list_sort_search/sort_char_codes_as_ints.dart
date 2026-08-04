// vybe-test: dart/list_sort_search/sort_char_codes_as_ints
// origin: languages/dart/tests/dart/test_list_sort_search.rs

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

void __vybeMain() {
  var list = ['z'.codeUnitAt(0), 'a'.codeUnitAt(0), 'm'.codeUnitAt(0)];
  list.sort();
  __p(String.fromCharCodes(list));
}

void main() {
  __vybeMain();
  __check('amz');
}
