// vybe-test: dart/string_interpolation/triple_quoted_preserves_embedded_single_quotes
// origin: languages/dart/tests/dart/test_string_interpolation.rs

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
  var text = '''it's fine''';
  __p(text.contains("'"));
}

void main() {
  __vybeMain();
  __check('true');
}
