// vybe-test: dart/list_core/list_insert_at_start_shifts_elements
// origin: languages/dart/tests/dart/test_list_core.rs

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
  var list = [2, 3];
  list.insert(0, 1);
  __p(list.first);
}

void main() {
  __vybeMain();
  __check('1');
}
