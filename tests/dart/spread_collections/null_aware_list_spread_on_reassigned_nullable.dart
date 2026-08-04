// vybe-test: dart/spread_collections/null_aware_list_spread_on_reassigned_nullable
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
  List<int>? data = null;
  var before = [...?data, 1];
  data = [2];
  var after = [...?data, 3];
  __p(before.join(','));
  __p(after.join(','));
}

void main() {
  __vybeMain();
  __check('1\n2,3');
}
