// vybe-test: dart/iterable_generate/iterable_generate_for_in_loop
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
  var sum = 0;
  for (var n in Iterable.generate(4, (i) => i + 1)) {
    sum = sum + n;
  }
  __p(sum);
}

void main() {
  __vybeMain();
  __check('10');
}
