// vybe-test: dart/typedefs_core/generic_typedef_string_map_getter
// origin: languages/dart/tests/dart/test_typedefs_core.rs

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

typedef Lookup = Map<String, String>;
String read(Lookup table, String key) {
  return table[key] ?? 'missing';
}
void __vybeMain() {
  Lookup labels = {'en': 'hello', 'fr': 'bonjour'};
  __p(read(labels, 'en'));
  __p(read(labels, 'de'));
}

void main() {
  __vybeMain();
  __check('hello\nmissing');
}
