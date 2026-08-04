// vybe-test: dart/custom_iterator/custom_iterator_string_values
// origin: languages/dart/tests/dart/test_custom_iterator.rs

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

class WordIterator implements Iterator<String> {
  final List<String> words;
  int _i = -1;
  WordIterator(this.words);
  @override
  String get current => words[_i];
  @override
  bool moveNext() {
    if (_i + 1 < words.length) {
      _i = _i + 1;
      return true;
    }
    return false;
  }
}
void __vybeMain() {
  var it = WordIterator(['a', 'b', 'c']);
  var text = '';
  while (it.moveNext()) {
    text = text + it.current;
  }
  __p(text);
}

void main() {
  __vybeMain();
  __check('abc');
}
