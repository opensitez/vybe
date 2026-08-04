// vybe-test: dart/typedefs_core/generic_typedef_map_alias_keys
// origin: languages/dart/tests/dart/test_typedefs_core.rs

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

typedef StringIntMap = Map<String, int>;
void __vybeMain() {
  StringIntMap scores = {'a': 1, 'b': 2};
  __p(scores.keys.length);
  __p(scores['b']);
}

void main() {
  __vybeMain();
  __check('2\n2');
}
