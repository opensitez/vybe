// vybe-test: dart/no_such_method/no_such_method_custom_to_string_still_works
// origin: languages/dart/tests/dart/test_no_such_method.rs

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

class Labeled {
  @override
  String toString() {
    return 'labeled';
  }
}
void __vybeMain() {
  __p(Labeled().toString());
}

void main() {
  __vybeMain();
  __check('labeled');
}
