// vybe-test: dart/iterable_methods/iterable_followed_by_then_to_list
// origin: languages/dart/tests/dart/test_iterable_methods.rs

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
  Iterable<int> a = [1];
  var list = a.followedBy([2, 3]).toList();
  __p(list.length);
  __p(list[2]);
}

void main() {
  __vybeMain();
  __check('3\n3');
}
