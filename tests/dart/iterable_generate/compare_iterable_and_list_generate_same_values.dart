// vybe-test: dart/iterable_generate/compare_iterable_and_list_generate_same_values
// origin: languages/dart/tests/dart/test_iterable_generate.rs

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
  var lazy = Iterable.generate(4, (i) => i * 2);
  var eager = List.generate(4, (i) => i * 2);
  __p(lazy.join(','));
  __p(eager.join(','));
}

void main() {
  __vybeMain();
  __check('0,2,4,6\n0,2,4,6');
}
