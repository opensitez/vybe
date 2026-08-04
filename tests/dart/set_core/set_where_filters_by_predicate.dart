// vybe-test: dart/set_core/set_where_filters_by_predicate
// origin: languages/dart/tests/dart/test_set_core.rs

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
  var s = {1, 2, 3, 4, 5};
  var evens = s.where((e) => e % 2 == 0).toList()..sort();
  __p(evens.join(','));
}

void main() {
  __vybeMain();
  __check('2,4');
}
