// vybe-test: dart/typedefs_core/typedef_named_target_function
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

typedef Formatter = String Function({required String prefix, required String body});
String format({required String prefix, required String body}) {
  return '$prefix:$body';
}
void __vybeMain() {
  Formatter fn = format;
  __p(fn(prefix: 'id', body: '42'));
}

void main() {
  __vybeMain();
  __check('id:42');
}
