// vybe-test: dart/exceptions_core/custom_exception_thrown_from_method
// origin: languages/dart/tests/dart/test_exceptions_core.rs

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

class ParseError implements Exception {
  final String detail;
  ParseError(this.detail);
}
class Parser {
  int parse(String s) {
    if (s.isEmpty) throw ParseError('empty');
    return int.parse(s);
  }
}
void __vybeMain() {
  try {
    Parser().parse('');
  } catch (e) {
    var err = e as ParseError;
    __p(err.detail);
  }
}

void main() {
  __vybeMain();
  __check('empty');
}
