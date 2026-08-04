// vybe-test: dart/list_sort_search/sort_pairs_by_second_component_via_wrapper
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
  var list = ['x:3', 'y:1', 'z:2'];
  list.sort((a, b) {
    var av = int.parse(a.split(':')[1]);
    var bv = int.parse(b.split(':')[1]);
    return av.compareTo(bv);
  });
  __p(list.join('|'));
}

void main() {
  __vybeMain();
  __check('y:1|z:2|x:3');
}
