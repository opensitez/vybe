// vybe-test: dart/cascades/cascade_string_substring_chain_preserves_full_length
// origin: languages/dart/tests/dart/test_cascades.rs

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
  var word = 'abcdef';
  word..substring(1, 3)..substring(4);
  __p(word.length);
  __p(word);
}

void main() {
  __vybeMain();
  __check('6\nabcdef');
}
