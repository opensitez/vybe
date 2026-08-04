// vybe-test: dart/abstract_members/abstract_method_returns_bool
// origin: languages/dart/tests/dart/test_abstract_members.rs

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

abstract class Check {
  bool ok();
}
class AlwaysYes extends Check {
  bool ok() {
    return true;
  }
}
void __vybeMain() {
  __p(AlwaysYes().ok());
}

void main() {
  __vybeMain();
  __check('true');
}
