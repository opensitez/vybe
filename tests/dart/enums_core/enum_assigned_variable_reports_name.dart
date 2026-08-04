// vybe-test: dart/enums_core/enum_assigned_variable_reports_name
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

enum Fruit { apple, banana, cherry }
void __vybeMain() {
  var f = Fruit.cherry;
  __p(f.name);
}

void main() {
  __vybeMain();
  __check('cherry');
}
