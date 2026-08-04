// vybe-test: dart/iterable_generate/iterable_generate_modulo_pattern
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
  var seq = Iterable.generate(6, (i) => i % 3);
  __p(seq.join(','));
}

void main() {
  __vybeMain();
  __check('0,1,2,0,1,2');
}
