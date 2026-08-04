// vybe-test: dart/getters_setters/getter_string_representation_from_fields
// origin: languages/dart/tests/dart/test_getters_setters.rs

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

class User {
  String first = 'Ada';
  String last = 'Lovelace';
  String get fullName {
    return first + ' ' + last;
  }
}
void __vybeMain() {
  __p(User().fullName);
}

void main() {
  __vybeMain();
  __check('Ada Lovelace');
}
