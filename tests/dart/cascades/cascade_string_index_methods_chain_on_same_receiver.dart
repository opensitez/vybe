// vybe-test: dart/cascades/cascade_string_index_methods_chain_on_same_receiver
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
  var word = 'ababa';
  word..indexOf('b')..lastIndexOf('b');
  __p(word.indexOf('b'));
  __p(word.lastIndexOf('b'));
}

void main() {
  __vybeMain();
  __check('1\n3');
}
