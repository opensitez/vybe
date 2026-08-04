// vybe-test: dart/callable_objects/call_with_one_optional_positional
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

class Join {
  String call([String a = 'x', String b = 'y']) {
    return a + b;
  }
}
void __vybeMain() {
  __p(Join()('a'));
}

void main() {
  __vybeMain();
  __check('ay');
}
