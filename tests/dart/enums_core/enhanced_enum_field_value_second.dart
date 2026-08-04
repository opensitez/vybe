// vybe-test: dart/enums_core/enhanced_enum_field_value_second
// origin: languages/dart/tests/dart/test_enums_core.rs

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

enum HttpCode {
  ok(200),
  notFound(404);
  final int code;
  const HttpCode(this.code);
}
void __vybeMain() {
  __p(HttpCode.notFound.code);
}

void main() {
  __vybeMain();
  __check('404');
}
