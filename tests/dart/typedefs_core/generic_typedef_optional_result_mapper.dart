// vybe-test: dart/typedefs_core/generic_typedef_optional_result_mapper
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

typedef Parser<T> = T? Function(String);
int? parseInt(String raw) {
  if (raw == '42') {
    return 42;
  }
  return null;
}
void __vybeMain() {
  Parser<int> parse = parseInt;
  __p(parse('42'));
  __p(parse('nope') == null);
}

void main() {
  __vybeMain();
  __check('42\ntrue');
}
