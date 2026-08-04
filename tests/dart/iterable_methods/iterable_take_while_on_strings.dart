// vybe-test: dart/iterable_methods/iterable_take_while_on_strings
// origin: languages/dart/tests/dart/test_iterable_methods.rs

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
  Iterable<String> words = ['a', 'ab', 'abc', 'b'];
  __p(words.takeWhile((w) => w.startsWith('a')).join('|'));
}

void main() {
  __vybeMain();
  __check('a|ab|abc');
}
