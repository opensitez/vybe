// vybe-test: dart/callable_objects/call_recursive_through_field
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

class Rec {
  int depth;
  Rec(this.depth);
  int call(int n) {
    if (depth <= 0) {
      return n;
    }
    return n + Rec(depth - 1)(n - 1);
  }
}
void __vybeMain() {
  __p(Rec(2)(3));
}

void main() {
  __vybeMain();
  __check('6');
}
