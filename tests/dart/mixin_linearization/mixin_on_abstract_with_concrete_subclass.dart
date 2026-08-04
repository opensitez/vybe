// vybe-test: dart/mixin_linearization/mixin_on_abstract_with_concrete_subclass
// origin: languages/dart/tests/dart/test_mixin_linearization.rs

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

abstract class Repo {
  int size();
}
mixin Cache on Repo {
  String label() {
    return 'cached';
  }
}
class ListRepo extends Repo with Cache {
  int size() {
    return 5;
  }
}
void __vybeMain() {
  __p(ListRepo().size());
}

void main() {
  __vybeMain();
  __check('5');
}
