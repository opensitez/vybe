// vybe-test: dart/generics_core/typed_map_string_to_int
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
  Map<String, int> scores = {'Ada': 90, 'Bob': 85};
  __p(scores['Ada']);
  __p(scores.length);
}

void main() {
  __vybeMain();
  __check('90\n2');
}
