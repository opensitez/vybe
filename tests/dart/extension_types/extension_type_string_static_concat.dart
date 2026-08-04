// vybe-test: dart/extension_types/extension_type_string_static_concat
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

extension type Note(String text) {
  static Note join(Note a, Note b) {
    return Note(a.text + b.text);
  }
}
void __vybeMain() {
  Note n = Note.join(Note('a'), Note('b'));
  __p(n.text);
}

void main() {
  __vybeMain();
  __check('ab');
}
