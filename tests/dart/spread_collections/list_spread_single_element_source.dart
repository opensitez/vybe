// vybe-test: dart/spread_collections/list_spread_single_element_source
// origin: languages/dart/tests/dart/test_spread_collections.rs

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
  var one = [99];
  var out = [...one, 100];
  __p(out.length);
  __p(out[0]);
  __p(out[1]);
}

void main() {
  __vybeMain();
  __check('2\n99\n100');
}
