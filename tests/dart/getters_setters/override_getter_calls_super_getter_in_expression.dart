// vybe-test: dart/getters_setters/override_getter_calls_super_getter_in_expression
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

class Base {
  String get label {
    return 'base';
  }
}
class Derived extends Base {
  String get label {
    return super.label + '-ext';
  }
}
void __vybeMain() {
  __p(Derived().label);
}

void main() {
  __vybeMain();
  __check('base-ext');
}
