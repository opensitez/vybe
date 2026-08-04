// vybe-test: dart/enums_core/enum_member_from_values_by_index
// origin: languages/dart/tests/dart/test_enums_core.rs

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

enum Level { low, mid, high }
void __vybeMain() {
  var picked = Level.values[1];
  __p(picked.name);
}

void main() {
  __vybeMain();
  __check('mid');
}
