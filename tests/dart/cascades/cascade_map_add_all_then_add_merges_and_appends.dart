// vybe-test: dart/cascades/cascade_map_add_all_then_add_merges_and_appends
// origin: languages/dart/tests/dart/test_cascades.rs

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
  var scores = <String, int>{'a': 1};
  scores..addAll({'b': 2, 'c': 3})..add('d', 4);
  __p(scores.length);
  __p(scores['d']);
}

void main() {
  __vybeMain();
  __check('4\n4');
}
