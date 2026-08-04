// vybe-test: dart/getters_setters/computed_getter_sums_field_values
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

class Range {
  int start = 2;
  int end = 8;
  int get span {
    return end - start;
  }
}
void __vybeMain() {
  __p(Range().span);
}

void main() {
  __vybeMain();
  __check('6');
}
