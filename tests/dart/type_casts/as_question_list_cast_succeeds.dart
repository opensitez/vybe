// vybe-test: dart/type_casts/as_question_list_cast_succeeds
// origin: languages/dart/tests/dart/test_type_casts.rs

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
  Object value = [10, 20];
  var list = value as List?;
  __p(list?.length);
}

void main() {
  __vybeMain();
  __check('2');
}
