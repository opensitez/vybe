// vybe-test: dart/runes_code_units/string_length_vs_runes_mixed_ascii_emoji
// origin: languages/dart/tests/dart/test_runes_code_units.rs

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
  var s = 'a🙂b';
  __p(s.length);
  __p(s.runes.length);
}

void main() {
  __vybeMain();
  __check('4\n3');
}
