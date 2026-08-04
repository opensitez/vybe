// vybe-test: dart/getters_setters/simple_getter_reads_backing_field
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

class Counter {
  int _count = 0;
  int get count {
    return _count;
  }
}
void __vybeMain() {
  __p(Counter().count);
}

void main() {
  __vybeMain();
  __check('0');
}
