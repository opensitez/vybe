// vybe-test: dart/extension_types/extension_type_string_contains_substring
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

extension type Path(String value) {
  bool containsPart(String part) {
    return value.contains(part);
  }
}
void __vybeMain() {
  Path p = Path('/usr/bin');
  __p(p.containsPart('bin'));
}

void main() {
  __vybeMain();
  __check('true');
}
