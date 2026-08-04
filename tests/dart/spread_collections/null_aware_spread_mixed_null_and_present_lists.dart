// vybe-test: dart/spread_collections/null_aware_spread_mixed_null_and_present_lists
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
  List<int>? b = [2, 3];
  var out = [1, ...?a, ...?b, 4];
  __p(out.join(','));
}

void main() {
  __vybeMain();
  __check('1,2,3,4');
}
