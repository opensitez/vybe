// vybe-test: dart/enums_core/enum_values_map_to_names_joined
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

enum Axis { x, y, z }
void __vybeMain() {
  var names = Axis.values.map((d) => d.name).join('-');
  __p(names);
}

void main() {
  __vybeMain();
  __check('x-y-z');
}
