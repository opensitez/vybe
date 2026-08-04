// vybe-test: dart/classes_core/class_method_returning_string_concatenation
// origin: languages/dart/tests/dart/test_classes_core.rs

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

class Name {
  String first = 'Ada';
  String last = 'Lovelace';
  String full() {
    return first + ' ' + last;
  }
}
void __vybeMain() {
  __p(Name().full());
}

void main() {
  __vybeMain();
  __check('Ada Lovelace');
}
