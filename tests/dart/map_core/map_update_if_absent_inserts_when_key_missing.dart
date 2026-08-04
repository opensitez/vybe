// vybe-test: dart/map_core/map_update_if_absent_inserts_when_key_missing
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
  var m = <String, int>{};
  m.update('created', (v) => v + 1, ifAbsent: () => 0);
  __p(m['created']);
}

void main() {
  __vybeMain();
  __check('0');
}
