// vybe-test: dart/abstract_members/abstract_method_null_return
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

abstract class Finder {
  String? find(int id);
}
class NullFinder extends Finder {
  String? find(int id) {
    return null;
  }
}
void __vybeMain() {
  __p(NullFinder().find(1) == null);
}

void main() {
  __vybeMain();
  __check('true');
}
