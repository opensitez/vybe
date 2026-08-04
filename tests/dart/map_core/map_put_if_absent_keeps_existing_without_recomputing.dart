// vybe-test: dart/map_core/map_put_if_absent_keeps_existing_without_recomputing
// origin: languages/dart/tests/dart/test_map_core.rs

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
  var m = {'keep': 11};
  var v = m.putIfAbsent('keep', () => 99);
  __p(v);
  __p(m['keep']);
}

void main() {
  __vybeMain();
  __check('11\n11');
}
