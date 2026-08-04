// vybe-test: dart/spread_collections/null_aware_spread_both_null_adds_only_literals
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
  List<int>? a = null;
  List<int>? b = null;
  var out = [...?a, ...?b, 9];
  __p(out.length);
  __p(out[0]);
}

void main() {
  __vybeMain();
  __check('1\n9');
}
