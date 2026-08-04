// vybe-test: dart/extension_types/extension_type_string_char_at_index
// origin: languages/dart/tests/dart/test_extension_types.rs

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

extension type Word(String value) {
  String charAt(int index) {
    return value[index];
  }
}
void __vybeMain() {
  Word w = Word('dart');
  __p(w.charAt(0));
  __p(w.charAt(3));
}

void main() {
  __vybeMain();
  __check('d\nt');
}
