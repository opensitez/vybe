// vybe-test: dart/typedefs_core/generic_typedef_callback_with_type_param
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

typedef Callback<T> = void Function(T);
var seen = <String>[];
void remember(String value) {
  seen.add(value);
}
void __vybeMain() {
  Callback<String> cb = remember;
  cb('dart');
  __p(seen.join(','));
}

void main() {
  __vybeMain();
  __check('dart');
}
