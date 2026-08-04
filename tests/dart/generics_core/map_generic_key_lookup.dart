// vybe-test: dart/generics_core/map_generic_key_lookup
// origin: languages/dart/tests/dart/test_generics_core.rs

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
  Map<int, String> table = {1: 'one', 2: 'two'};
  __p(table[2]);
  __p(table.containsKey(1));
}

void main() {
  __vybeMain();
  __check('two\ntrue');
}
