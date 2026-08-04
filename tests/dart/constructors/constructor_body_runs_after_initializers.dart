// vybe-test: dart/constructors/constructor_body_runs_after_initializers
// origin: languages/dart/tests/dart/test_constructors.rs

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

class Log {
  int step;
  Log(int s) : step = s {
    step = step + 10;
  }
}
void __vybeMain() {
  __p(Log(1).step);
}

void main() {
  __vybeMain();
  __check('11');
}
