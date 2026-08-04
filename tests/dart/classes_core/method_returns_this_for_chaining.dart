// vybe-test: dart/classes_core/method_returns_this_for_chaining
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

class Builder {
  int v = 0;
  Builder add(int n) {
    v = v + n;
    return this;
  }
}
void __vybeMain() {
  var b = Builder();
  b.add(2).add(3);
  __p(b.v);
}

void main() {
  __vybeMain();
  __check('5');
}
