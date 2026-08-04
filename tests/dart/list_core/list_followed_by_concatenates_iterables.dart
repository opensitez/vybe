// vybe-test: dart/list_core/list_followed_by_concatenates_iterables
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
  var first = [1, 2];
  var combined = first.followedBy([3, 4]);
  __p(combined.join(','));
}

void main() {
  __vybeMain();
  __check('1,2,3,4');
}
