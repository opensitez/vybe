// vybe-test: dart/callable_objects/call_multiple_instances_independent
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

class Box {
  int scale;
  Box(this.scale);
  int call(int n) {
    return n * scale;
  }
}
void __vybeMain() {
  __p(Box(2)(3) + Box(5)(3));
}

void main() {
  __vybeMain();
  __check('21');
}
