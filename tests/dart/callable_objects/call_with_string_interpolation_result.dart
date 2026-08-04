// vybe-test: dart/callable_objects/call_with_string_interpolation_result
// origin: languages/dart/tests/dart/test_callable_objects.rs

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

class Tag {
  String call(String name, int id) {
    return '$name#$id';
  }
}
void __vybeMain() {
  __p(Tag()('item', 42));
}

void main() {
  __vybeMain();
  __check('item#42');
}
