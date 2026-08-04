// vybe-test: dart/comparable_ordering/comparable_sort_strings_alphabetically
// origin: languages/dart/tests/dart/test_comparable_ordering.rs

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

class Word implements Comparable<Word> {
  String text;
  Word(this.text);
  int compareTo(Word other) => text.compareTo(other.text);
}
void __vybeMain() {
  var words = [Word('cherry'), Word('apple'), Word('banana')];
  words.sort();
  __p(words[0].text);
  __p(words[2].text);
}

void main() {
  __vybeMain();
  __check('apple\ncherry');
}
