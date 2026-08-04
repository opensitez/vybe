// vybe-test: dart/abstract_members/abstract_method_finds_value
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
class MapFinder extends Finder {
  String? find(int id) {
    if (id == 1) {
      return 'found';
    }
    return null;
  }
}
void __vybeMain() {
  __p(MapFinder().find(1));
}

void main() {
  __vybeMain();
  __check('found');
}
