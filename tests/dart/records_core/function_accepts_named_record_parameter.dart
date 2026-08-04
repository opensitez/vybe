// vybe-test: dart/records_core/function_accepts_named_record_parameter
// origin: languages/dart/tests/dart/test_records_core.rs

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

String formatUser(({String name, int id}) user) {
  return user.name + ':' + user.id.toString();
}
void __vybeMain() {
  __p(formatUser((name: 'Eve', id: 7)));
}

void main() {
  __vybeMain();
  __check('Eve:7');
}
